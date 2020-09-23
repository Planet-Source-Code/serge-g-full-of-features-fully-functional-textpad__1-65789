VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTextPad 
   Caption         =   "TextPad"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9630
   Icon            =   "frmTextPad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4425
      Top             =   4245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":10C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1374
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1628
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7035
      Top             =   5505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":1FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":256A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":2E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":2F9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":30F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":3692
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":37EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":40C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":49A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":4F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":54D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":562E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":5788
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":5D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":5E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":5FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":6570
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":66CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":6C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":71FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":7798
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":7D32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbox1 
      Height          =   2550
      Left            =   15
      TabIndex        =   8
      Top             =   780
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   4498
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmTextPad.frx":82CC
   End
   Begin TabDlg.SSTab ssTab1 
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   1376
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Text   "
      TabPicture(0)   =   "frmTextPad.frx":834E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tbr1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmTextPad.frx":84A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optMC"
      Tab(1).Control(1)=   "optww"
      Tab(1).Control(2)=   "optFromT"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "tpReplace"
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(6)=   "tpSearch"
      Tab(1).Control(7)=   "Line4"
      Tab(1).Control(8)=   "Line3"
      Tab(1).Control(9)=   "Line2"
      Tab(1).Control(10)=   "Line1"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Tools"
      TabPicture(2)   =   "frmTextPad.frx":8602
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmb1"
      Tab(2).Control(1)=   "cmdApplyFont"
      Tab(2).Control(2)=   "tbr2"
      Tab(2).ControlCount=   3
      Begin VB.ComboBox cmb1 
         Height          =   315
         Left            =   -71310
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   375
         Width           =   2895
      End
      Begin VB.CommandButton cmdApplyFont 
         Caption         =   "Apply"
         Height          =   300
         Left            =   -68325
         TabIndex        =   14
         Top             =   390
         Width           =   1095
      End
      Begin VB.CheckBox optMC 
         Caption         =   "Match Case"
         Height          =   270
         Left            =   -66735
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   405
         Width           =   1095
      End
      Begin VB.CheckBox optww 
         Caption         =   "Whole Word"
         Height          =   275
         Left            =   -67860
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   405
         Width           =   1095
      End
      Begin VB.CheckBox optFromT 
         Caption         =   "From Top"
         Height          =   275
         Left            =   -68985
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   405
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar tbr1 
         Height          =   330
         Left            =   75
         TabIndex        =   9
         Top             =   375
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   27
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open File"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tbrNew"
               Object.ToolTipText     =   "New..."
               ImageIndex      =   8
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "tbrSameItm"
                     Text            =   "Same Pad"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "tbrNewItm"
                     Text            =   "New Pad"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save your work"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print your work"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "spell"
               Object.ToolTipText     =   "Spell Check"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy Text"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut Text"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste Text"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "delete"
               Object.ToolTipText     =   "Delete Selected Text"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "clipboard"
               Object.ToolTipText     =   "View Clipboard"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo Action"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo Action"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Selected Text to Bold"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "Selected Text to Italic"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline Selected Text "
               ImageIndex      =   15
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strikeout Selected Text"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Search and Replace"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ucase"
               Object.ToolTipText     =   "Convert text to upper case"
               ImageIndex      =   24
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ucaseAll"
                     Text            =   "All Text"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "ucaseSeled"
                     Text            =   "Selected Text"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "lcase"
               Object.ToolTipText     =   "Conver text to lower case"
               ImageIndex      =   25
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "lcaseAll"
                     Text            =   "All Text"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "lcaseSeled"
                     Text            =   "Selected Text"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Replace"
         Height          =   255
         Left            =   -72015
         TabIndex        =   7
         Top             =   435
         Width           =   855
      End
      Begin VB.TextBox tpReplace 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71115
         TabIndex        =   6
         Top             =   420
         Width           =   1965
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   270
         Left            =   -74925
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox tpSearch 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   -74160
         TabIndex        =   4
         Top             =   420
         Width           =   1980
      End
      Begin MSComctlLib.Toolbar tbr2 
         Height          =   330
         Left            =   -74925
         TabIndex        =   13
         Top             =   360
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "calendar"
               Object.ToolTipText     =   "See the Calendar"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "timer"
               Object.ToolTipText     =   "Set Alarm / Timer"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "calc"
               Object.ToolTipText     =   "Load Calculator"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "lock"
               Object.ToolTipText     =   "Lock the Text"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "unlock"
               Object.ToolTipText     =   "Unlock"
               ImageIndex      =   8
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "options"
               Object.ToolTipText     =   "Change Settings..."
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "font"
               Object.ToolTipText     =   "Change FontName"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "about"
               Object.ToolTipText     =   "About"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               Object.ToolTipText     =   "See Help"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         X1              =   -69075
         X2              =   -69075
         Y1              =   285
         Y2              =   765
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   -69060
         X2              =   -69060
         Y1              =   300
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         X1              =   -72105
         X2              =   -72105
         Y1              =   300
         Y2              =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   -72090
         X2              =   -72090
         Y1              =   315
         Y2              =   720
      End
   End
   Begin MSComDlg.CommonDialog cmnDlg2 
      Left            =   5880
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Locate the *.Dat file"
   End
   Begin MSComDlg.CommonDialog cmnDlg1 
      Left            =   6480
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTBUndo2 
      Height          =   855
      Left            =   1800
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmTextPad.frx":875C
   End
   Begin RichTextLib.RichTextBox RTBUndo 
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmTextPad.frx":87DE
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6075
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Dockable"
            TextSave        =   "Dockable"
            Object.Tag             =   "panelOnTop"
            Object.ToolTipText     =   "Click to be always on top"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "3:22 AM"
            Object.Tag             =   "panelTime"
            Object.ToolTipText     =   "Click to hide time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "7/23/2006"
            Object.Tag             =   "panelDate"
            Object.ToolTipText     =   "Click to hide date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1014
            MinWidth        =   1014
            TextSave        =   "CAPS"
            Object.Tag             =   "panelCaps"
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.Tag             =   "Num Lock"
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   882
            Text            =   "Saved"
            TextSave        =   "Saved"
            Object.Tag             =   "fileInfo"
            Object.ToolTipText     =   "Current file information"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1005
            MinWidth        =   441
            Text            =   "Line: 1"
            TextSave        =   "Line: 1"
            Object.Tag             =   "lineNumber"
            Object.ToolTipText     =   "Line number"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   556
            MinWidth        =   556
            Picture         =   "frmTextPad.frx":8860
            Key             =   "bold"
            Object.ToolTipText     =   "Font weight Bold"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   556
            MinWidth        =   556
            Picture         =   "frmTextPad.frx":8CA2
            Key             =   "italic"
            Object.ToolTipText     =   "Font weight Italic"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   556
            MinWidth        =   556
            Picture         =   "frmTextPad.frx":90E4
            Key             =   "underline"
            Object.ToolTipText     =   "Font wieght Underline"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   556
            MinWidth        =   556
            Picture         =   "frmTextPad.frx":9526
            Key             =   "strike"
            Object.ToolTipText     =   "Font wieght Strikethru"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuNewItem 
            Caption         =   "New (Same Pad)"
         End
         Begin VB.Menu mnuNewPadItem 
            Caption         =   "New (New Pad)"
         End
      End
      Begin VB.Menu mnuOpenItem 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSaveItem 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAsItem 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSep40 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintItem 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndoItem 
         Caption         =   "Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRedoItem 
         Caption         =   "Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Begin VB.Menu mnuCutItem 
            Caption         =   "Cut                        Ctrl + X"
         End
         Begin VB.Menu mnuPvtCutItem 
            Caption         =   "Private Cut           Shft + Ctrl + X"
         End
         Begin VB.Menu mnuSpCutItem 
            Caption         =   "Special Cut           Alt + Ctrl + X"
         End
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Begin VB.Menu mnuCopyItem 
            Caption         =   "Copy                      Ctrl + C"
         End
         Begin VB.Menu mnuPrvtCopyItem 
            Caption         =   "Private Copy         Shft + Ctrl + C"
         End
         Begin VB.Menu mnuSCopyItem 
            Caption         =   "Special Copy         Alt + Ctrl + C"
         End
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Begin VB.Menu mnuPasteItem 
            Caption         =   "Paste                      Ctrl + V"
         End
         Begin VB.Menu mnuPrvtPasteItem 
            Caption         =   "Private Paste         Shft + Ctrl + C"
         End
      End
      Begin VB.Menu mnuCBViewItem 
         Caption         =   "View Clipboard"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindItem 
         Caption         =   "Find    F4"
      End
      Begin VB.Menu mnuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChkSpellItem 
         Caption         =   "Check Spelling (MSWord)"
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInTimeItem 
         Caption         =   "Insert Time"
      End
      Begin VB.Menu mnuInDateItem 
         Caption         =   "Insert Date"
      End
      Begin VB.Menu mnuCustomInsert 
         Caption         =   "Custom Insert"
         Begin VB.Menu mnuWord1Item 
            Caption         =   "Word 1"
         End
         Begin VB.Menu mnuWord2Item 
            Caption         =   "Word 2"
         End
         Begin VB.Menu mnuWord3Item 
            Caption         =   "Word 3"
         End
         Begin VB.Menu mnuWord4Item 
            Caption         =   "Word 4"
         End
         Begin VB.Menu mnuWord5Item 
            Caption         =   "Word 5"
         End
      End
      Begin VB.Menu mnuHdnSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAllItem 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndentAll 
         Caption         =   "Indent All Text"
         Begin VB.Menu mnuIndenAllLeftItem 
            Caption         =   "Indent Left"
         End
         Begin VB.Menu mnuUnindentAllLeftItem 
            Caption         =   "Unindent Left"
         End
         Begin VB.Menu mnuSep21 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndentAllRightItem 
            Caption         =   "Indent All Right"
         End
         Begin VB.Menu mnuUnindentAllRightItem 
            Caption         =   "Unindent Right"
         End
         Begin VB.Menu mnuSep19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndentAllBothItem 
            Caption         =   "Indent Left And Right"
         End
         Begin VB.Menu mnuUnindentAllBothItem 
            Caption         =   "Unindent Left And Right"
         End
         Begin VB.Menu mnuSep25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBullInAll 
            Caption         =   "Bullet Indent"
         End
         Begin VB.Menu mnuUndoBullInAll 
            Caption         =   "Undo Bullet Indent"
         End
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent Selected Text"
         Enabled         =   0   'False
         Begin VB.Menu mnuIndentStandardItem 
            Caption         =   "Indent"
         End
         Begin VB.Menu mnuUndoIndentItem 
            Caption         =   "Undo Indent"
         End
         Begin VB.Menu mnuSep17 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndentRightItem 
            Caption         =   "Indent Right"
         End
         Begin VB.Menu mnuUndoIndentRightItem 
            Caption         =   "Undo Indent Right"
         End
         Begin VB.Menu mnuSep18 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndentBothItem 
            Caption         =   "Indent Left And Right"
         End
         Begin VB.Menu mnuUndoIndentBothItem 
            Caption         =   "Undo Indent"
         End
         Begin VB.Menu mnuSep16 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndentItem 
            Caption         =   "Bullet Indent"
         End
         Begin VB.Menu mnuUndoBullItem 
            Caption         =   "Undo Bullet Indent"
         End
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnseledTxtItem2 
         Caption         =   "Unselect Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuShowSeledItem 
         Caption         =   "> No Text Selected <"
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuDelRegItem 
         Caption         =   "Delete All Settings"
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockItem 
         Caption         =   "Lock"
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetTmrItem 
         Caption         =   "Set Timer..."
      End
      Begin VB.Menu mnuSeeCalenItem 
         Caption         =   "Show Calendar..."
      End
      Begin VB.Menu mnuSep35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalcItem 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingItem 
         Caption         =   "Settings..."
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "F&ont"
      Begin VB.Menu mnuBold 
         Caption         =   "Bold"
         Begin VB.Menu mnuBoldItem 
            Caption         =   "Selected Text"
         End
         Begin VB.Menu mnuBoldAllItem 
            Caption         =   "All Text"
         End
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "Italic"
         Begin VB.Menu mnuItalicItem 
            Caption         =   "Selected Text"
         End
         Begin VB.Menu mnuItalicAllItem 
            Caption         =   "All Text"
         End
      End
      Begin VB.Menu mnuULAll 
         Caption         =   "Underline"
         Begin VB.Menu mnuULItem 
            Caption         =   "Selected Text"
         End
         Begin VB.Menu mnuULAllItem 
            Caption         =   "All Text"
         End
      End
      Begin VB.Menu mnuStrike 
         Caption         =   "StrikeThru"
         Begin VB.Menu mnuSTItem 
            Caption         =   "Selected Text"
         End
         Begin VB.Menu mnuStrikeAllItem 
            Caption         =   "All Text"
         End
      End
      Begin VB.Menu mnuRegText 
         Caption         =   "Regular"
         Begin VB.Menu mnuRegularFontItem 
            Caption         =   "Selected Text"
         End
         Begin VB.Menu mnuRegAllItem 
            Caption         =   "All Text"
         End
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu subMnuFont 
         Caption         =   "Font"
         Begin VB.Menu mnuFontNameItem 
            Caption         =   "Font Face..."
         End
         Begin VB.Menu mnuFontSizeItem 
            Caption         =   "Font Size..."
         End
         Begin VB.Menu mnuFColorItem 
            Caption         =   "Font Color..."
         End
         Begin VB.Menu mnuSep26 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQFontItem 
            Caption         =   "Quick Font  "
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuSep12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "Zoom"
            Begin VB.Menu mnuXSmallItem 
               Caption         =   "Extra Small"
            End
            Begin VB.Menu mnuSmallItem 
               Caption         =   "Small"
            End
            Begin VB.Menu mnuMediumItem 
               Caption         =   "Medium"
            End
            Begin VB.Menu mnuLargeItem 
               Caption         =   "Large"
            End
            Begin VB.Menu mnuXLargeItem 
               Caption         =   "Extra Large"
            End
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnselectItem 
         Caption         =   "Unselect Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeledText 
         Caption         =   "> No Text Selected <"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSCItem 
         Caption         =   "Shortcut Reference"
      End
      Begin VB.Menu mnuAllHelpItem 
         Caption         =   "General Help"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "HiddenMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuHdnSaveItem 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHdnCopyItem 
         Caption         =   "Copy                      Ctrl + C"
      End
      Begin VB.Menu mnuHdnCutItem 
         Caption         =   "Cut                         Ctrl + X"
      End
      Begin VB.Menu mnuHdnSCopyItem 
         Caption         =   "Special Copy          Alt + Ctrl + C"
      End
      Begin VB.Menu mnuHdnSCutItem 
         Caption         =   "Special Cut             Alt + Ctrl + X"
      End
      Begin VB.Menu mnuHdnPasteItem 
         Caption         =   "Paste                     Ctrl + V"
      End
      Begin VB.Menu mnuHdnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHdnPCopyItem 
         Caption         =   "Private Copy         Shft + Ctrl + C"
      End
      Begin VB.Menu mnuHdnPCutItem 
         Caption         =   "Private Cut            Shft + Ctrl + X"
      End
      Begin VB.Menu mnuHdnPPasteItem 
         Caption         =   "Private Paste        Shft + Ctrl + V"
      End
      Begin VB.Menu mnuHdnSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHdnDelItem 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuHdnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHdnSelAllItem 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnuHidden2 
      Caption         =   "HiddenMenu2"
      Visible         =   0   'False
      Begin VB.Menu mnuHideClockItem 
         Caption         =   "Hide Clock"
      End
      Begin VB.Menu mnuSetTimerItem 
         Caption         =   "Set Timer"
      End
   End
   Begin VB.Menu mnuHiddenCal 
      Caption         =   "HiddenMenuCalendar"
      Visible         =   0   'False
      Begin VB.Menu mnuHideDateItem 
         Caption         =   "Hide Date"
      End
      Begin VB.Menu mnuShowCalItem 
         Caption         =   "See Calendar"
      End
   End
   Begin VB.Menu mnuHdnYesNo 
      Caption         =   "yesNo"
      Visible         =   0   'False
      Begin VB.Menu mnuYesNoOkItem 
         Caption         =   "Confirm"
      End
      Begin VB.Menu mnuYesNoNoItem 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmTextPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delSet As Boolean
Dim isBold As Boolean
Dim isItalic As Boolean
Dim isUL As Boolean
Dim isST As Boolean
Dim savedIt As Boolean
Dim saveString As String
Dim firstSave As Boolean
Dim comeBack As Integer
Dim saveExt As String
Dim openNew As Boolean
Dim beingOpened As Boolean
Dim saveFilePath As String
Dim txtFileChars As Double
Dim pvtCopy As Variant
Dim doSpecialCopy As Boolean
Dim totalChars As Variant
Dim textSelected As Boolean
Dim lastChar As Boolean

Private Sub cmdApplyFont_Click()

    PopupMenu mnuHdnYesNo
    rtbox1.SetFocus

End Sub

Private Sub Command1_Click()

    mnuFindItem_Click
    frmSearchTxt.txtFind.Text = tpSearch.Text
    
    If optFromT.Value = 1 Then
        frmSearchTxt.chkStartFromTop.Value = 1
    Else
        frmSearchTxt.chkStartFromTop.Value = 0
    End If
    
    If optww.Value = 1 Then
        frmSearchTxt.chkWWord.Value = 1
    Else
        frmSearchTxt.chkWWord.Value = 0
    End If
    
    If optMC.Value = 1 Then
        frmSearchTxt.chkMCase.Value = 1
    Else
        frmSearchTxt.chkMCase.Value = 0
    End If

End Sub

Private Sub Command2_Click()

    If tpReplace.Text <> "" Then
        
        mnuFindItem_Click
        frmSearchTxt.txtFind.Text = tpSearch.Text
        frmSearchTxt.txtReplace.Enabled = True
        frmSearchTxt.txtReplace.BackColor = QBColor(15)
        frmSearchTxt.txtReplace.Text = tpReplace.Text
        frmSearchTxt.cmdReplace.Enabled = True
        frmSearchTxt.cmdReplaceAll.Enabled = True
    
        If optFromT.Value = 1 Then
            frmSearchTxt.chkStartFromTop.Value = 1
        Else
            frmSearchTxt.chkStartFromTop.Value = 0
        End If
    
        If optww.Value = 1 Then
            frmSearchTxt.chkWWord.Value = 1
        Else
            frmSearchTxt.chkWWord.Value = 0
        End If
     
        If optMC.Value = 1 Then
            frmSearchTxt.chkMCase.Value = 1
        Else
            frmSearchTxt.chkMCase.Value = 0
        End If
     
        End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Dim strLoc As String
    
    strLoc = Command

    delSet = False
    isBold = False
    isItalic = False
    isUL = False
    isST = False
    savedIt = True
    firstSave = True
    openNew = False
    beingOpened = False
    doSpecialCopy = False
    textSelected = False
    isTimerSet = False
    pvtCopy = ""
    saveExt = ".txt"
    comeBack = 0
    Form_Resize
    getRegSettings
    saveFileLocation
    Me.Height = currHt
    Me.Width = currWd
    Me.Top = currTop
    Me.Left = currLft
    setStatBar
    SetTopMost
    anyOpenFiles
    mySavedOptions

        If openNew = False Then
            writeTempReg
        End If
    
    rtbox1.BackColor = defaultBGColor

100
    If strLoc = "" Then
        Me.Caption = "Untitled" & openFileCount & ".txt"
        saveString = "Untitled" & openFileCount
        
        rtbox1.BulletIndent = bullIndentBy
        rtbox1.Font.Name = defaultFont
          rtbox1.SelStart = 0
          rtbox1.SelLength = Len(rtbox1.Text)
          rtbox1.SelColor = defaultFontColor
    
        Select Case defaultIndentAll
        Case "Indent Left"
            With rtbox1
                .SelStart = 1
                .SelLength = Len(rtbox1.Text)
                .SelIndent = indentLeft
            End With
        Case "Indent Right"
             With rtbox1
                .SelStart = 1
                .SelLength = Len(rtbox1.Text)
                .SelRightIndent = indentRight
             End With
        Case "Indent Left and Right"
            With rtbox1
                .SelStart = 1
                .SelLength = Len(rtbox1.Text)
                .SelIndent = indentLeft
                .SelRightIndent = indentRight
            End With
        Case "Bullet Indent"
            rtbox1.SelStart = 0
            rtbox1.SelLength = Len(rtbox1.Text)
            rtbox1.SelBullet = True
        Case "No Indent"
             With rtbox1
                 .SelStart = 1
                 .SelLength = Len(rtbox1.Text)
                 .SelIndent = 0
                 .SelRightIndent = 0
              End With
    
              rtbox1.SelStart = 0
              rtbox1.SelLength = Len(rtbox1.Text)
              rtbox1.SelBullet = False
        End Select

    Else
        On Error Resume Next
        strLoc = Mid(strLoc, 2, Len(strLoc) - 2)
        If FileLen(strLoc) > 2000000 Then
            cnt = MsgBox("The file size exeeds the limitation of the TextPad. The program may not load properly. Continue?", vbYesNo, "Warning")
            If cnt = vbNo Then
                strLoc = ""
                GoTo 100
            End If
        
        End If
        
            rtbox1.LoadFile strLoc
            s = InStrRev(strLoc, "\", -1, vbTextCompare)
            saveString = Mid(strLoc, s + 1, 100)
            Me.Caption = saveString '& " | Path : " & strLoc
        
    End If
    Me.Show
    rtbox1.SetFocus
    
    If Len(rtbox1.Text) = 0 Then
        lastChar = True
    End If
    
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuUndoItem.Enabled = False
    mnuRedoItem.Enabled = False
    
    tbr1.Buttons(9).Enabled = False
    tbr1.Buttons(8).Enabled = False
    tbr1.Buttons(16).Enabled = False
    
    'Hidden Menu
    
    mnuHdnCopyItem.Enabled = False
    mnuHdnCutItem.Enabled = False
    mnuHdnSCopyItem.Enabled = False
    mnuHdnSCutItem.Enabled = False
    mnuHdnPCopyItem.Enabled = False
    mnuHdnPCutItem.Enabled = False
    mnuHdnDelItem.Enabled = False
    
    totalChars = Len(rtbox1.Text)
        
    setFileInfo (10)
    setLineNum
      
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    customWordMenu
    
    Load frmTextPad_ImageList
    
    For q = 0 To Screen.FontCount - 1
        cmb1.AddItem Screen.Fonts(q)
    Next q
    
    cmb1.Text = rtbox1.Font.Name
    
   ' RTBox1.Text
   
   If rtbox1.SelBold = True Then
       StatusBar1.Panels(8).Bevel = sbrInset
       StatusBar1.Panels(8).Key = "bold_down"
       StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
   Else
       StatusBar1.Panels(8).Bevel = sbrRaised
       StatusBar1.Panels(8).Key = "bold"
       StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
   End If
   
   If rtbox1.SelItalic = True Then
       StatusBar1.Panels(9).Bevel = sbrInset
       StatusBar1.Panels(9).Key = "italic_down"
       StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
   Else
       StatusBar1.Panels(9).Bevel = sbrRaised
       StatusBar1.Panels(9).Key = "italic"
       StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
   End If
   
   If rtbox1.SelUnderline = True Then
       StatusBar1.Panels(10).Bevel = sbrInset
       StatusBar1.Panels(10).Key = "underline_down"
       StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
   Else
       StatusBar1.Panels(10).Bevel = sbrRaised
       StatusBar1.Panels(10).Key = "underline"
       StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
   End If
   
   If rtbox1.SelStrikeThru = True Then
       StatusBar1.Panels(11).Bevel = sbrInset
       StatusBar1.Panels(11).Key = "strike_down"
       StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
   Else
       StatusBar1.Panels(11).Bevel = sbrRaised
       StatusBar1.Panels(11).Key = "strike"
       StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
   End If
               
End Sub

Private Sub Form_Resize()

    On Error GoTo errhand
    
    If Me.WindowState = vbMinimized Then
        
    ElseIf Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        
    End If
    
    If Me.Height < 950 Then
        Me.Height = 950
    End If
    If Me.Width < 1850 Then
        Me.Width = 1850
    End If

    ssTab1.Width = Me.Width - 100

    rtbox1.Width = Me.Width - 95
    rtbox1.Height = Me.Height - ssTab1.Height - StatusBar1.Height - 675
    'rtbox1.Height = Me.Height - ssTab1.Height - 950
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''                      For                           ''
    ''                  Toolbar use                       ''
    ''                     Only                           ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'RTBox1.Height = RTBox1.Height - 435
    'RTBox1.Top = 435

errhand:
    If Err.Number <> 0 Then
        'MsgBox (Err.Description & " Error Number :" & Err.Number)
        Resume Next
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    Dim theTemp
    Dim t As Boolean
    
    t = TopMost
    TopMost = False
    SetTopMost
    
    If savedIt = False Then
    
        'qt = MsgBox("Do you want to save " & saveString & "? ", vbYesNoCancel, "Save " & saveString & " ?")
        Load frmMsgBoxYNC
        frmMsgBoxYNC.Caption = "TextPad"
        
        If Len(saveString) > 30 Then
            theTemp = Mid(saveString, Len(saveString) - 25, Len(saveString))
            theTemp = "...\..." & theTemp
        Else
            theTemp = saveString
        End If
        
        frmMsgBoxYNC.lblTop = "The text in the " & theTemp & " has changed."
        frmMsgBoxYNC.lblBottom = "Do you want to save the changes?"
        frmMsgBoxYNC.Show vbModal
        qt = quitConfirm
        If qt = vbYes Then
            TopMost = t
            remSettings
            deleteTempReg
            
            If firstSave = False Then
                Save_It
            Else
                save_It_As
            End If
                
                Unload frmOnTop
                End
            
        ElseIf qt = vbNo Then
            TopMost = t
            remSettings
            deleteTempReg
            Unload frmOnTop
            End
        ElseIf qt = vbCancel Then
            Cancel = -1
            TopMost = t
            SetTopMost
        End If
        
    Else
        deleteTempReg
        TopMost = t
        remSettings
        Unload frmOnTop
        End
    End If

End Sub

Private Sub mnuAboutItem_Click()

    Dim t As Boolean
    
    t = TopMost
    TopMost = False
    SetTopMost

    frmAbout.Show vbModal, Me
    
    TopMost = t
    SetTopMost

End Sub

Private Sub mnuAllHelpItem_Click()

    frmHelp.Show vbModeless, Me

End Sub

Private Sub mnuBoldAllItem_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelBold = True
    rtbox1.SelStart = Len(rtbox1.Text)
    rtbox1.SelLength = 0
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuBoldItem_Click()

'    If RTBox1.SelBold = True Then
'        MsgBox ("B")
'    ElseIf RTBox1.SelBold = False Then
'        MsgBox ("NB")
'    Else
'        MsgBox ("BOTH")
'    End If
'
'    If isBold = False Then
'        RTBox1.SelBold = True
'        isBold = True
'    Else
'        RTBox1.SelBold = False
'        isBold = False
'    End If

    If rtbox1.SelLength > 0 Then
        rtbox1.SelBold = True
        savedIt = False
        setFileInfo (10)
    End If
    
    mnuUnselectItem_Click

End Sub

Private Sub mnuBullInAll_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelBullet = True
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuCalcItem_Click()

    frmCalc.Show

End Sub

Private Sub mnuCBViewItem_Click()

    Dim t As Boolean
    
    t = TopMost
    TopMost = False
    SetTopMost

    publicCopy = pvtCopy
    Load frmClipBoard
    frmClipBoard.Show vbModal
    pvtCopy = publicCopy
    
    TopMost = t
    SetTopMost

End Sub

Private Sub mnuChkSpellItem_Click()

    Load frmSpellCheck
    frmSpellCheck.Show vbModal, Me
    

'
'    Dim tempStat As Boolean
'
'    tempStat = TopMost
'    TopMost = False
'    SetTopMost
'
'    On Error GoTo errhand
'
'    Dim X As Object
'    Set X = CreateObject("word.application")
'    X.Visible = False
'    X.Documents.Add
'    X.Selection.Text = rtbox1.Text
'    X.ActiveDocument.CheckSpelling
'    rtbox1.Text = X.Selection.Text
'    X.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
'    X.Quit
'
'    Set X = Nothing
'
'    MsgBox ("Done checking spelling"), , "Text Pad"
'
'    'setPause (2)
'
'    TopMost = tempStat
'    SetTopMost
'
'errhand:
'If Err.Number = 0 Or Err.Number = 20 Then
'    Resume Next
'Else
'    MsgBox ("Unable to perform spell check")
'    Exit Sub
'End If
'
End Sub

Private Sub mnuCopyItem_Click()

    Clipboard.Clear
    Clipboard.SetText rtbox1.SelText
    rtbox1.SetFocus

End Sub

Private Sub mnuCutItem_Click()

    Clipboard.Clear
    Clipboard.SetText rtbox1.SelText
    rtbox1.SelText = ""
    rtbox1.SetFocus
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuDelRegItem_Click()

    mnuDelRegItem.Checked = Not mnuDelRegItem.Checked
    If mnuDelRegItem.Checked = True Then
        delSet = True
    Else
        delSet = False
    End If

End Sub

Private Sub mnuEdit_Click()

    If Clipboard.GetText = "" Then
        mnuPasteItem.Enabled = False
    Else
        mnuPasteItem.Enabled = True
    End If
    
    If pvtCopy = "" Then
        mnuPrvtPasteItem.Enabled = False
    Else
        mnuPrvtPasteItem.Enabled = True
    End If

End Sub

Private Sub mnuExitItem_Click()

    Unload Me
    
End Sub

Private Sub mnuFColorItem_Click()

    cmnDlg1.Flags = cdlCCFullOpen Or cdlCCRGBInit
    cmnDlg1.ShowColor
    rtbox1.SelColor = cmnDlg1.Color

End Sub

Private Sub mnuFindItem_Click()

    Load frmSearchTxt
    frmSearchTxt.txtFind.Text = rtbox1.SelText
    rtbox1.SelLength = 0
    rtbox1.SetFocus
    frmSearchTxt.Show vbModeless, Me
    'searchOnTop

End Sub

Private Sub mnuFontNameItem_Click()

    cmnDlg1.Flags = cdlCFBoth Or cdlCFApply Or cdlCFEffects
    cmnDlg1.ShowFont
    rtbox1.SelFontName = cmnDlg1.FontName
    rtbox1.SelFontSize = cmnDlg1.FontSize
    
    If cmnDlg1.FontBold = True Then
        rtbox1.SelBold = True
    ElseIf cmnDlg1.FontBold = False Then
        rtbox1.SelBold = False
    End If
    
    If cmnDlg1.FontItalic = True Then
        rtbox1.SelItalic = True
    ElseIf cmnDlg1.FontItalic = False Then
        rtbox1.SelItalic = False
    End If
    
    If cmnDlg1.FontUnderline = True Then
        rtbox1.SelUnderline = True
    ElseIf cmnDlg1.FontUnderline = False Then
        rtbox1.SelUnderline = False
    End If
    
    If cmnDlg1.FontStrikethru = True Then
        rtbox1.SelStrikeThru = True
    ElseIf cmnDlg1.FontStrikethru = False Then
        rtbox1.SelStrikeThru = False
    End If
    
    rtbox1.SelFontSize = cmnDlg1.FontSize
    
End Sub

Private Sub mnuFontSizeItem_Click()

    Dim t As Boolean
    
    t = TopMost
    TopMost = False
    SetTopMost

    currFontSize = rtbox1.Font.Size
    currFontName = rtbox1.Font.Name
    frmFontSize.Show vbModal, Me
    rtbox1.SelFontSize = currFontSize
    
    TopMost = t
    SetTopMost

End Sub

Private Sub mnuHdnCopyItem_Click()

    mnuCopyItem_Click

End Sub

Private Sub mnuHdnCutItem_Click()

    mnuCutItem_Click

End Sub

Private Sub mnuHdnDelItem_Click()

    rtbox1.SelText = ""
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuHdnPasteItem_Click()

    mnuPasteItem_Click

End Sub

Private Sub mnuHdnPCopyItem_Click()

    mnuPrvtCopyItem_Click

End Sub

Private Sub mnuHdnPCutItem_Click()

    mnuPvtCutItem_Click

End Sub

Private Sub mnuHdnPPasteItem_Click()

    mnuPrvtPasteItem_Click

End Sub

Private Sub mnuHdnSaveItem_Click()

    mnuSaveItem_Click

End Sub

Private Sub mnuHdnSCopyItem_Click()

    mnuSCopyItem_Click

End Sub

Private Sub mnuHdnSCutItem_Click()

    mnuSpCutItem_Click

End Sub

Private Sub mnuHdnSelAllItem_Click()

    mnuSelAllItem_Click

End Sub



Private Sub mnuHelpSCItem_Click()
    
    frmQHelp.Show vbModeless, Me

End Sub

Private Sub mnuHidden_Click()

    If Clipboard.GetText = "" Then
        mnuHdnPasteItem.Enabled = False
    Else
        mnuHdnPasteItem.Enabled = True
    End If
    
    If pvtCopy = "" Then
        mnuHdnPPasteItem.Enabled = False
    Else
        mnuHdnPPasteItem.Enabled = True
    End If

End Sub

Private Sub mnuHideClockItem_Click()

    If showTime = True Then
       StatusBar1.Panels(2).Bevel = sbrNoBevel
       StatusBar1.Panels(2).Style = sbrText
       StatusBar1.Panels(2).Text = ""
       StatusBar1.Panels(2).ToolTipText = "Click to show time"
       showTime = False
    Else
       StatusBar1.Panels(2).Bevel = sbrInset
       StatusBar1.Panels(2).Style = sbrTime
       StatusBar1.Panels(2).ToolTipText = "Click to hide time"
       showTime = True
    End If

End Sub

Private Sub mnuHideDateItem_Click()

    If showDate = True Then
        StatusBar1.Panels(3).Bevel = sbrNoBevel
        StatusBar1.Panels(3).Style = sbrText
        StatusBar1.Panels(3).ToolTipText = "Click to show date"
        showDate = False
    Else
        StatusBar1.Panels(3).Bevel = sbrInset
        StatusBar1.Panels(3).Style = sbrDate
        StatusBar1.Panels(3).ToolTipText = "Click to hide date"
        showDate = True
    End If

End Sub

Public Sub mnuInDateItem_Click()

    rtbox1.SelText = Date
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndenAllLeftItem_Click()

    With rtbox1
      .SelStart = 1
      .SelLength = Len(rtbox1.Text)
      .SelIndent = indentLeft
    End With
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndentAllBothItem_Click()

    With rtbox1
      .SelStart = 1
      .SelLength = Len(rtbox1.Text)
      .SelIndent = indentLeft
      .SelRightIndent = indentRight
    End With
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndentAllRightItem_Click()

    With rtbox1
      .SelStart = 1
      .SelLength = Len(rtbox1.Text)
      .SelRightIndent = indentRight
    End With
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
        
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndentBothItem_Click()

    With rtbox1
       .SelIndent = indentLeft
       .SelRightIndent = indentRight
    End With
   
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
       
   savedIt = False
   setFileInfo (10)

End Sub

Private Sub mnuIndentItem_Click()

    rtbox1.SelBullet = True
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndentRightItem_Click()

    rtbox1.SelRightIndent = indentRight
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuIndentStandardItem_Click()

    rtbox1.SelIndent = indentLeft
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)
    
End Sub

Public Sub mnuInTimeItem_Click()

    rtbox1.SelText = Time
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuItalicAllItem_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelItalic = True
    rtbox1.SelStart = Len(rtbox1.Text)
    rtbox1.SelLength = 0
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuItalicItem_Click()

    If rtbox1.SelLength > 0 Then
        rtbox1.SelItalic = True
        savedIt = False
        setFileInfo (10)
    End If

    mnuUnselectItem_Click

End Sub

Private Sub mnuLargeItem_Click()

    rtbox1.Font.Size = 13
    currViewSize = 13
    uncheckViewSize
    mnuLargeItem.Checked = True

End Sub

Private Sub mnuLockItem_Click()

    rtbox1.Locked = Not rtbox1.Locked
    
    If rtbox1.Locked = True Then
        mnuLockItem.Checked = True
        tbr2.Buttons(5).Value = tbrPressed
        tbr2.Buttons(6).Value = tbrUnpressed
    Else
        mnuLockItem.Checked = False
        tbr2.Buttons(5).Value = tbrUnpressed
        tbr2.Buttons(6).Value = tbrPressed
    End If

End Sub

Private Sub mnuMediumItem_Click()

    rtbox1.Font.Size = 11
    currViewSize = 11
    uncheckViewSize
    mnuMediumItem.Checked = True

End Sub

Private Sub mnuNewItem_Click()

    comeBack = 0
    If savedIt = True Then
        rtbox1.Text = ""
        RTBUndo.Text = ""
        RTBUndo2.Text = ""
        openNew = True
        deleteTempReg
        Form_Load
        Me.Caption = "Untitled" & openFileCount & ".txt"
        setFileInfo (0)
     Else
        comeBack = 2
        wantToSave
    End If

End Sub

Private Sub mnuNewPadItem_Click()

    On Error GoTo erHand
    
    Dim proID
    proID = Shell(exeLocation & ".exe", vbNormalFocus)
    
erHand:
Resume Next
'proID = Shell("C:\Programming and Scripting\Old Inetpub\samples\Learning_VisualBasic\MyLessons\TextPad\ExeTemp\text pad" & ".exe", vbNormalFocus)
'Exit Sub

End Sub

Private Sub mnuOpenItem_Click()

    On Error Resume Next

    comeBack = 0
    If savedIt = True Then
        beingOpened = True
        cmnDlg1.FileName = ""
        cmnDlg1.Filter = "Text Document (*.txt) |*.txt|Rich Text Document (*.rtf )|*.rtf|All Files|*.*"
        cmnDlg1.ShowOpen
        
        tempS = cmnDlg1.FileName
        If tempS <> "" Then
            
            If FileLen(cmnDlg1.FileName) > 2000000 Then
                cnt = MsgBox("The file size exeeds the limitation of the TextPad. The program may not load properly. Continue?", vbYesNo, "Warning")
                
                If cnt = vbNo Then Exit Sub
                
            End If
            
                saveString = cmnDlg1.FileName
                rtbox1.LoadFile saveString
                txtFileChars = Len(rtbox1.Text)
                rtbox1.SelStart = txtFileChars
                        
        End If
                        
        If tempS <> "" Then
            Me.Caption = cmnDlg1.FileTitle '+ " | Path: " + saveString
            saveExt = InStrRev(saveString, ".", -1, vbTextCompare)
            saveExt = Mid(saveString, saveExt + 1, 100)
            saveExt = "." & saveExt
            firstSave = False
            setFileInfo (10)
            totalChars = Len(rtbox1.Text)
            setLineNum
            
        Else
            'Nothing
        End If
    Else
        comeBack = 1
        wantToSave
        'mnuSaveItem_Click
    End If
    
End Sub

Private Sub mnuPasteItem_Click()

    rtbox1.SelText = Clipboard.GetText
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuPrintItem_Click()
 
    cmnDlg1.ShowPrinter

End Sub

Private Sub mnuPrvtCopyItem_Click()

    pvtCopy = ""
    pvtCopy = rtbox1.SelText

End Sub

Private Sub mnuPrvtPasteItem_Click()

    rtbox1.SelText = pvtCopy
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuPvtCutItem_Click()

    pvtCopy = ""
    pvtCopy = rtbox1.SelText
    rtbox1.SelText = ""
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuQFontItem_Click()

     frmQFont.Show vbModeless, Me

End Sub

Private Sub mnuRedoItem_Click()

    rtbox1.Text = RTBUndo.Text
    mnuUndoItem.Enabled = True
    tbr1.Buttons(15).Enabled = True
    
    mnuRedoItem.Enabled = False
    tbr1.Buttons(16).Enabled = False
    rtbox1.SelStart = Len(rtbox1.Text)
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuRegAllItem_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelBold = False
    rtbox1.SelItalic = False
    rtbox1.SelUnderline = False
    rtbox1.SelStrikeThru = False
    rtbox1.SelStart = Len(rtbox1.Text)
    rtbox1.SelLength = 0
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuRegularFontItem_Click()

    If rtbox1.SelLength > 0 Then
        rtbox1.SelBold = False
        rtbox1.SelItalic = False
        rtbox1.SelUnderline = False
        rtbox1.SelStrikeThru = False
        savedIt = False
        setFileInfo (10)
    End If

    mnuUnselectItem_Click

End Sub

Private Sub mnuSaveAsItem_Click()
    
    On Error GoTo errhand
    
    cmnDlg1.FileName = saveString
    cmnDlg1.CancelError = True
    cmnDlg1.Filter = "Text Document (*.txt)|*.txt|Rich Text Document (*.rtf) |*.rtf|All Files|*.*"
    cmnDlg1.Flags = cdlOFNOverwritePrompt
    cmnDlg1.ShowSave
    saveFilePath = cmnDlg1.FileName
    saveString = cmnDlg1.FileTitle
    
    If saveString <> "" Then
        rtbox1.SaveFile saveString, rtfText
        Me.Caption = saveString '& " | Path: " & saveFilePath
        firstSave = False
        savedIt = True
        beingOpened = True
        setFileInfo (10)
    End If
    comeBackWhere

errhand:
    If Err.Number = 32755 Then
        comeBack = 0
        Exit Sub
    Else
        Resume Next
    End If

End Sub

Private Sub mnuSaveItem_Click()

    If firstSave = False Then
        rtbox1.SaveFile saveString
        savedIt = True
        beingOpened = True
        setFileInfo (10)
    Else
        mnuSaveAsItem_Click
    End If
    
    comeBackWhere

End Sub

Private Sub mnuSCopyItem_Click()

    Dim tempStr As Variant
    tempStr = ""
    
    If doSpecialCopy = False Then
        Clipboard.Clear
        doSpecialCopy = True
    End If
    
    tempStr = Clipboard.GetText
    Clipboard.Clear
    Clipboard.SetText tempStr & specialCopyChar & rtbox1.SelText

End Sub

Private Sub mnuSeeCalenItem_Click()

    mnuShowCalItem_Click

End Sub

Private Sub mnuSelAllItem_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SetFocus

End Sub

Private Sub mnuSetTimerItem_Click()

    frmSetTimer.Show vbModeless, Me
    frmSetTimer.WindowState = vbNormal

End Sub

Private Sub mnuSettingItem_Click()
    
    frmOptions.Show vbModeless, Me
    
End Sub

Private Sub mnuSetTmrItem_Click()

    mnuSetTimerItem_Click

End Sub

Private Sub mnuShowCalItem_Click()

    frmCalendar.Show vbModeless, Me

End Sub

Private Sub mnuSmallItem_Click()

    rtbox1.Font.Size = 9
    currViewSize = 9
    uncheckViewSize
    mnuSmallItem.Checked = True

End Sub

Private Sub mnuSpCutItem_Click()

    Dim tempStr As Variant
    tempStr = ""
    
    If doSpecialCopy = False Then
        Clipboard.Clear
        doSpecialCopy = True
    End If
    
    tempStr = Clipboard.GetText
    Clipboard.Clear
    Clipboard.SetText tempStr & specialCopyChar & rtbox1.SelText
    rtbox1.SelText = ""

    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuSTItem_Click()

    If rtbox1.SelLength > 0 Then
        rtbox1.SelStrikeThru = True
        savedIt = False
        setFileInfo (10)
    End If

    mnuUnselectItem_Click

End Sub

Private Sub mnuStrikeAllItem_Click()
    
    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelStrikeThru = True
    rtbox1.SelStart = Len(rtbox1.Text)
    rtbox1.SelLength = 0
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuULAllItem_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelUnderline = True
    rtbox1.SelStart = Len(rtbox1.Text)
    rtbox1.SelLength = 0
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuULItem_Click()

    If rtbox1.SelLength > 0 Then
        rtbox1.SelUnderline = True
        savedIt = False
        setFileInfo (10)
    End If

    mnuUnselectItem_Click

End Sub

Private Sub mnuUndoBullInAll_Click()

    rtbox1.SelStart = 0
    rtbox1.SelLength = Len(rtbox1.Text)
    rtbox1.SelBullet = False
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
        
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUndoBullItem_Click()

    rtbox1.SelBullet = False
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUndoIndentBothItem_Click()

    rtbox1.SelIndent = 0
    rtbox1.SelRightIndent = 0
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUndoIndentItem_Click()

    rtbox1.SelIndent = 0
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUndoIndentRightItem_Click()

    rtbox1.SelRightIndent = 0
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUndoItem_Click()

    rtbox1.Text = RTBUndo.Text
    mnuUndoItem.Enabled = False
    tbr1.Buttons(15).Enabled = False
    
    mnuRedoItem.Enabled = True
    tbr1.Buttons(16).Enabled = True
    rtbox1.SelStart = Len(rtbox1.Text)
    savedIt = False
    setFileInfo (10)
    
End Sub

Private Sub mnuUnindentAllBothItem_Click()

    With rtbox1
       .SelStart = 1
       .SelLength = Len(rtbox1.Text)
       .SelIndent = 0
       .SelRightIndent = 0
    End With
   
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUnindentAllLeftItem_Click()

    With rtbox1
      .SelStart = 1
      .SelLength = Len(rtbox1.Text)
      .SelIndent = 0
    End With
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUnindentAllRightItem_Click()

    With rtbox1
      .SelStart = 1
      .SelLength = Len(rtbox1.Text)
      .SelRightIndent = 0
    End With
    
    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
        
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuUnseledTxtItem2_click()

    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0

End Sub

Private Sub mnuUnselectItem_Click()

    rtbox1.SelStart = rtbox1.SelStart + Len(rtbox1.SelText)
    rtbox1.SelLength = 0
    
End Sub

Private Sub mnuWord1Item_Click()

    rtbox1.SelText = customWord(1)
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuWord2Item_Click()

    rtbox1.SelText = customWord(2)
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuWord3Item_Click()

    rtbox1.SelText = customWord(3)
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuWord4Item_Click()

    rtbox1.SelText = customWord(4)
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuWord5Item_Click()

    rtbox1.SelText = customWord(5)
    
    savedIt = False
    setFileInfo (10)

End Sub

Private Sub mnuXLargeItem_Click()

    rtbox1.Font.Size = 16
    currViewSize = 16
    uncheckViewSize
    mnuXLargeItem.Checked = True

End Sub

Private Sub mnuXSmallItem_Click()

    rtbox1.Font.Size = 8
    currViewSize = 8
    uncheckViewSize
    mnuXSmallItem.Checked = True
    
End Sub

Private Sub mnuYesNoNoItem_Click()

    cmb1.Text = rtbox1.Font.Name
    rtbox1.SetFocus

End Sub

Private Sub mnuYesNoOkItem_Click()

    rtbox1.SelFontName = cmb1.Text
    rtbox1.SetFocus

End Sub

Private Sub RTBox1_Change()
    
    On Error Resume Next
    
    setLineNum
    RTBUndo.Text = RTBUndo2.Text
    RTBUndo2.Text = rtbox1.Text
    mnuUndoItem.Enabled = True
    tbr1.Buttons(15).Enabled = True

'    If beingOpened = False Then
'        savedIt = False
'    Else
'        beingOpened = False
'    End If

End Sub

Private Sub RTBox1_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    If isQFActive = True Then
'        Unload frmQFont
'    End If
    
    setLineNum
   
End Sub

Private Sub RTBox1_DblClick()

    Unload frmQFont

End Sub

Private Sub RTBox1_KeyDown(KeyCode As Integer, Shift As Integer)

    
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
       
    If textSelected = True Then
    
        If ShiftDown And CtrlDown And KeyCode = 88 Then
           mnuPvtCutItem_Click
           KeyCode = 0
        ElseIf ShiftDown And CtrlDown And KeyCode = 67 Then
           mnuPrvtCopyItem_Click
           KeyCode = 0
        ElseIf AltDown And CtrlDown And KeyCode = 88 Then
            mnuSpCutItem_Click
            KeyCode = 0
        ElseIf AltDown And CtrlDown And KeyCode = 67 Then
            mnuSCopyItem_Click
            KeyCode = 0
        End If
    End If
    
    If mnuUndoItem.Enabled = True Then
        If CtrlDown And KeyCode = 90 Then
            mnuUndoItem_Click
            KeyCode = 0
        End If
    End If
    
    If mnuRedoItem.Enabled = True Then
        If CtrlDown And KeyCode = 89 Then
            mnuRedoItem_Click
            KeyCode = 0
        End If
    End If
        
    If ShiftDown And CtrlDown And KeyCode = 86 Then
       mnuPrvtPasteItem_Click
       KeyCode = 0
    ElseIf CtrlDown And KeyCode = 70 Then
        mnuFindItem_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF4 Then
        mnuFindItem_Click
        KeyCode = 0
    ElseIf CtrlDown And AltDown And KeyCode = 72 Then
        frmHelp.Show vbModal, Me
        KeyCode = 0
    ElseIf KeyCode = vbKeyF1 Then
        frmHelp.Show vbModal, Me
        KeyCode = 0
    ElseIf CtrlDown And KeyCode = 83 Then
        'save
        KeyCode = 0
    ElseIf AltDown And CtrlDown And KeyCode = 84 Then
        mnuInTimeItem_Click
        KeyCode = 0
    ElseIf AltDown And CtrlDown And KeyCode = 68 Then
        mnuInDateItem_Click
        KeyCode = 0
    ElseIf ShiftDown And AltDown And KeyCode = 66 Then
        mnuBoldItem_Click
        KeyCode = 0
    ElseIf ShiftDown And AltDown And KeyCode = 73 Then
        mnuItalicItem_Click
        KeyCode = 0
    ElseIf ShiftDown And AltDown And KeyCode = 85 Then
        mnuULItem_Click
        KeyCode = 0
    ElseIf ShiftDown And AltDown And KeyCode = 83 Then
        mnuSTItem_Click
        KeyCode = 0
    ElseIf CtrlDown And KeyCode = 81 Then
        frmQFont.Show vbModeless, Me
        KeyCode = 0
    End If
      
    If KeyCode = vbKeyF2 Then   ' Display help
        frmQHelp.Show vbModeless, Me
    End If
   
    'MsgBox (KeyCode)
    
    'MsgBox (totalChars)
    'MsgBox (KeyCode)
    'MsgBox (totalChars)
    
    If StatusBar1.Panels(8).Key = "bold" And rtbox1.SelBold = True Then
        StatusBar1.Panels(8).Key = "bold_down"
        StatusBar1.Panels(8).Bevel = sbrInset
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
    ElseIf StatusBar1.Panels(8).Key = "bold_down" And rtbox1.SelBold = False Then
        StatusBar1.Panels(8).Key = "bold"
        StatusBar1.Panels(8).Bevel = sbrRaised
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
    ElseIf StatusBar1.Panels(8).Key = "bold_mixed" And rtbox1.SelBold = False Then
        StatusBar1.Panels(8).Key = "bold"
        StatusBar1.Panels(8).Bevel = sbrRaised
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
    ElseIf StatusBar1.Panels(8).Key = "bold_mixed" And rtbox1.SelBold = True Then
        StatusBar1.Panels(8).Key = "bold_down"
        StatusBar1.Panels(8).Bevel = sbrInset
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
    End If
    
    If StatusBar1.Panels(9).Key = "italic" And rtbox1.SelItalic = True Then
        StatusBar1.Panels(9).Key = "italic_down"
        StatusBar1.Panels(9).Bevel = sbrInset
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
    ElseIf StatusBar1.Panels(9).Key = "italic_down" And rtbox1.SelItalic = False Then
        StatusBar1.Panels(9).Key = "italic"
        StatusBar1.Panels(9).Bevel = sbrRaised
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
    ElseIf StatusBar1.Panels(9).Key = "italic_mixed" And rtbox1.SelItalic = False Then
        StatusBar1.Panels(9).Key = "italic"
        StatusBar1.Panels(9).Bevel = sbrRaised
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
    ElseIf StatusBar1.Panels(9).Key = "italic_mixed" And rtbox1.SelItalic = True Then
        StatusBar1.Panels(9).Key = "italic_down"
        StatusBar1.Panels(9).Bevel = sbrInset
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
    End If
    
    If StatusBar1.Panels(10).Key = "underline" And rtbox1.SelUnderline = True Then
        StatusBar1.Panels(10).Key = "underline_down"
        StatusBar1.Panels(10).Bevel = sbrInset
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
    ElseIf StatusBar1.Panels(10).Key = "underline_down" And rtbox1.SelUnderline = False Then
        StatusBar1.Panels(10).Key = "underline"
        StatusBar1.Panels(10).Bevel = sbrRaised
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
    ElseIf StatusBar1.Panels(10).Key = "underline_mixed" And rtbox1.SelUnderline = False Then
        StatusBar1.Panels(10).Key = "underline"
        StatusBar1.Panels(10).Bevel = sbrRaised
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
    ElseIf StatusBar1.Panels(10).Key = "underline_mixed" And rtbox1.SelUnderline = True Then
        StatusBar1.Panels(10).Key = "underline_down"
        StatusBar1.Panels(10).Bevel = sbrInset
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
    End If
    
    If StatusBar1.Panels(11).Key = "strike" And rtbox1.SelStrikeThru = True Then
        StatusBar1.Panels(11).Key = "strike_down"
        StatusBar1.Panels(11).Bevel = sbrInset
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
    ElseIf StatusBar1.Panels(11).Key = "strike_down" And rtbox1.SelStrikeThru = False Then
        StatusBar1.Panels(11).Key = "strike"
        StatusBar1.Panels(11).Bevel = sbrRaised
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
    ElseIf StatusBar1.Panels(11).Key = "strike_mixed" And rtbox1.SelStrikeThru = False Then
        StatusBar1.Panels(11).Key = "strike"
        StatusBar1.Panels(11).Bevel = sbrRaised
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
    ElseIf StatusBar1.Panels(11).Key = "strike_mixed" And rtbox1.SelStrikeThru = True Then
        StatusBar1.Panels(11).Key = "strike_down"
        StatusBar1.Panels(11).Bevel = sbrInset
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
    End If
    
    cmb1.Text = rtbox1.SelFontName
    
End Sub

Private Sub RTBox1_KeyPress(KeyAscii As Integer)

    setLineNum
'    MsgBox (KeyAscii)
    
    If useLastChar = True Then
        
        If lastChar = True Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If KeyAscii = 46 Or KeyAscii = 33 Or KeyAscii = 63 Then
            lastChar = True
        ElseIf KeyAscii <> 32 Then
            lastChar = False
        End If
        
    End If
    
End Sub

Private Sub RTBox1_KeyUp(KeyCode As Integer, Shift As Integer)

    If totalChars <> Len(rtbox1.Text) Then
        savedIt = False
        setFileInfo (10)
    End If
    totalChars = Len(rtbox1.Text)
    setLineNum
    
  If StatusBar1.Panels(8).Key = "bold" And rtbox1.SelBold = True Then
        StatusBar1.Panels(8).Key = "bold_down"
        StatusBar1.Panels(8).Bevel = sbrInset
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
    ElseIf StatusBar1.Panels(8).Key = "bold_down" And rtbox1.SelBold = False Then
        StatusBar1.Panels(8).Key = "bold"
        StatusBar1.Panels(8).Bevel = sbrRaised
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
    ElseIf StatusBar1.Panels(8).Key = "bold_mixed" And rtbox1.SelBold = False Then
        StatusBar1.Panels(8).Key = "bold"
        StatusBar1.Panels(8).Bevel = sbrRaised
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
    ElseIf StatusBar1.Panels(8).Key = "bold_mixed" And rtbox1.SelBold = True Then
        StatusBar1.Panels(8).Key = "bold_down"
        StatusBar1.Panels(8).Bevel = sbrInset
        StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
    End If
    
    If StatusBar1.Panels(9).Key = "italic" And rtbox1.SelItalic = True Then
        StatusBar1.Panels(9).Key = "italic_down"
        StatusBar1.Panels(9).Bevel = sbrInset
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
    ElseIf StatusBar1.Panels(9).Key = "italic_down" And rtbox1.SelItalic = False Then
        StatusBar1.Panels(9).Key = "italic"
        StatusBar1.Panels(9).Bevel = sbrRaised
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
    ElseIf StatusBar1.Panels(9).Key = "italic_mixed" And rtbox1.SelItalic = False Then
        StatusBar1.Panels(9).Key = "italic"
        StatusBar1.Panels(9).Bevel = sbrRaised
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
    ElseIf StatusBar1.Panels(9).Key = "italic_mixed" And rtbox1.SelItalic = True Then
        StatusBar1.Panels(9).Key = "italic_down"
        StatusBar1.Panels(9).Bevel = sbrInset
        StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
    End If
    
    If StatusBar1.Panels(10).Key = "underline" And rtbox1.SelUnderline = True Then
        StatusBar1.Panels(10).Key = "underline_down"
        StatusBar1.Panels(10).Bevel = sbrInset
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
    ElseIf StatusBar1.Panels(10).Key = "underline_down" And rtbox1.SelUnderline = False Then
        StatusBar1.Panels(10).Key = "underline"
        StatusBar1.Panels(10).Bevel = sbrRaised
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
    ElseIf StatusBar1.Panels(10).Key = "underline_mixed" And rtbox1.SelUnderline = False Then
        StatusBar1.Panels(10).Key = "underline"
        StatusBar1.Panels(10).Bevel = sbrRaised
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
    ElseIf StatusBar1.Panels(10).Key = "underline_mixed" And rtbox1.SelUnderline = True Then
        StatusBar1.Panels(10).Key = "underline_down"
        StatusBar1.Panels(10).Bevel = sbrInset
        StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
    End If
    
    If StatusBar1.Panels(11).Key = "strike" And rtbox1.SelStrikeThru = True Then
        StatusBar1.Panels(11).Key = "strike_down"
        StatusBar1.Panels(11).Bevel = sbrInset
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
    ElseIf StatusBar1.Panels(11).Key = "strike_down" And rtbox1.SelStrikeThru = False Then
        StatusBar1.Panels(11).Key = "strike"
        StatusBar1.Panels(11).Bevel = sbrRaised
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
    ElseIf StatusBar1.Panels(11).Key = "strike_mixed" And rtbox1.SelStrikeThru = False Then
        StatusBar1.Panels(11).Key = "strike"
        StatusBar1.Panels(11).Bevel = sbrRaised
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
    ElseIf StatusBar1.Panels(11).Key = "strike_mixed" And rtbox1.SelStrikeThru = True Then
        StatusBar1.Panels(11).Key = "strike_down"
        StatusBar1.Panels(11).Bevel = sbrInset
        StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
    End If
    
End Sub

Private Sub RTBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuHidden
    End If

End Sub

Private Sub RTBox1_SelChange()

    On Error Resume Next
    
    Dim seledText As String
    
    textSelected = True
    mnuUnselectItem.Enabled = True
    mnuUnseledTxtItem2.Enabled = True
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
    tbr1.Buttons(8).Enabled = True
    tbr1.Buttons(9).Enabled = True
    tbr1.Buttons(25).ButtonMenus(2).Enabled = True
    tbr1.Buttons(26).ButtonMenus(2).Enabled = True
    
    'HiddenMenu (PopUpMenu)
    mnuHdnCopyItem.Enabled = True
    mnuHdnCutItem.Enabled = True
    mnuHdnSCopyItem.Enabled = True
    mnuHdnSCutItem.Enabled = True
    mnuHdnPCopyItem.Enabled = True
    mnuHdnPCutItem.Enabled = True
    mnuHdnDelItem.Enabled = True
        
    mnuIndent.Enabled = True
    seledText = Mid(rtbox1.SelText, 1, 15)
    
    If Len(seledText) >= 15 Then
        mnuSeledText.Caption = "> " & seledText & "... <"
        mnuShowSeledItem.Caption = "> " & seledText & "... <"
    Else
        mnuSeledText.Caption = "> " & seledText & " <"
        mnuShowSeledItem.Caption = "> " & seledText & " <"
    End If
     
    If rtbox1.SelLength = 0 Then
        mnuUnselectItem.Enabled = False
        mnuUnseledTxtItem2.Enabled = False
        mnuCut.Enabled = False
        mnuCopy.Enabled = False
        tbr1.Buttons(8).Enabled = False
        tbr1.Buttons(9).Enabled = False
        tbr1.Buttons(25).ButtonMenus(2).Enabled = False
        tbr1.Buttons(26).ButtonMenus(2).Enabled = False
        
        'Hidden Menu
        
        mnuHdnCopyItem.Enabled = False
        mnuHdnCutItem.Enabled = False
        mnuHdnSCopyItem.Enabled = False
        mnuHdnSCutItem.Enabled = False
        mnuHdnPCopyItem.Enabled = False
        mnuHdnPCutItem.Enabled = False
        mnuHdnDelItem.Enabled = False
        
        
        mnuIndent.Enabled = False
        textSelected = False
        mnuSeledText.Caption = "> No Text Selected <"
        mnuShowSeledItem.Caption = "> No Text Selected <"
    End If
    
    If rtbox1.SelLength > 0 Then
        
        If rtbox1.SelBold = True Then
            StatusBar1.Panels(8).Key = "bold_down"
            StatusBar1.Panels(8).Bevel = sbrInset
            StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Inset
        ElseIf rtbox1.SelBold = False Then
            StatusBar1.Panels(8).Key = "bold"
            StatusBar1.Panels(8).Bevel = sbrRaised
            StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Raised
        Else
            StatusBar1.Panels(8).Key = "bold_mixed"
            StatusBar1.Panels(8).Bevel = sbrNoBevel
            StatusBar1.Panels(8).Picture = frmTextPad_ImageList.img_Bold_Mixed
        End If
        
        If rtbox1.SelItalic = True Then
            StatusBar1.Panels(9).Key = "italic_down"
            StatusBar1.Panels(9).Bevel = sbrInset
            StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Inset
        ElseIf rtbox1.SelItalic = False Then
            StatusBar1.Panels(9).Key = "italic"
            StatusBar1.Panels(9).Bevel = sbrRaised
            StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Raised
        Else
            StatusBar1.Panels(9).Key = "italic_mixed"
            StatusBar1.Panels(9).Bevel = sbrNoBevel
            StatusBar1.Panels(9).Picture = frmTextPad_ImageList.img_Italic_Mixed
        End If
        
        If rtbox1.SelUnderline = True Then
            StatusBar1.Panels(10).Key = "underline_down"
            StatusBar1.Panels(10).Bevel = sbrInset
            StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Inset
        ElseIf rtbox1.SelUnderline = False Then
            StatusBar1.Panels(10).Key = "underline"
            StatusBar1.Panels(10).Bevel = sbrRaised
            StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Raised
        Else
            StatusBar1.Panels(10).Key = "underline_mixed"
            StatusBar1.Panels(10).Bevel = sbrNoBevel
            StatusBar1.Panels(10).Picture = frmTextPad_ImageList.img_Underline_Mixed
        End If
        
        If rtbox1.SelStrikeThru = True Then
            StatusBar1.Panels(11).Key = "strike_down"
            StatusBar1.Panels(11).Bevel = sbrInset
            StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Inset
        ElseIf rtbox1.SelStrikeThru = False Then
            StatusBar1.Panels(11).Key = "strike"
            StatusBar1.Panels(11).Bevel = sbrRaised
            StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Raised
        Else
            StatusBar1.Panels(11).Key = "strike_mixed"
            StatusBar1.Panels(11).Bevel = sbrNoBevel
            StatusBar1.Panels(11).Picture = frmTextPad_ImageList.img_Strike_Mixed
        End If
        
    End If

End Sub



Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'''''    If x > 0 And x < 1201 Then
'''''        MsgBox ("P1")
'''''    ElseIf x > 1244 And x < 2011 Then
'''''        MsgBox ("P2")
'''''    ElseIf x > 2054 And x < 3031 Then
'''''        MsgBox ("P3")
'''''    Else
'''''        MsgBox (x)
'''''    End If

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

    On Error Resume Next
    
    If Panel.Tag = "panelOnTop" Then
        If iAmActive = 10 Then
            Panel.ToolTipText = "Not available while searching"
            Panel.Text = "Not Available"
        Else
            If TopMost = False Then
                Panel.Bevel = sbrRaised
                Panel.Text = "Always on top"
                Panel.ToolTipText = "Click to make window dockable"
                Unload frmOnTop
                TopMost = True
                SetTopMost
            Else
                Panel.Bevel = sbrInset
                Panel.Text = "Dockable"
                Panel.ToolTipText = "Click to be always on top"
                TopMost = False
                SetTopMost
            End If
        End If
        
    ElseIf Panel.Tag = "panelTime" Then
        If isTimerSet = False Then
            mnuSetTimerItem.Caption = "Set Timer"
        ElseIf isTimerSet = True Then
            mnuSetTimerItem.Caption = "Stop Timer"
        End If
        If showTime = True Then
            mnuHideClockItem.Caption = "Hide Clock"
            PopupMenu mnuHidden2
        ElseIf showTime = False Then
            mnuHideClockItem.Caption = "Show Clock"
            PopupMenu mnuHidden2
        End If
        
'       If showTime = True Then
'           Panel.Bevel = sbrNoBevel
'           Panel.Style = sbrText
'           Panel.Text = ""
'           Panel.ToolTipText = "Click to show time"
'           showTime = False
'       Else
'           Panel.Bevel = sbrInset
'           Panel.Style = sbrTime
'           Panel.ToolTipText = "Click to hide time"
'           showTime = True
'       End If
            
    ElseIf Panel.Tag = "panelDate" Then
        
        If showDate = True Then
            mnuHideDateItem.Caption = "Hide Date"
            PopupMenu mnuHiddenCal
        ElseIf showDate = False Then
            mnuHideDateItem.Caption = "Show Date"
            PopupMenu mnuHiddenCal
        End If
'        If showDate = True Then
'            Panel.Bevel = sbrNoBevel
'            Panel.Style = sbrText
'            Panel.ToolTipText = "Click to show date"
'            showDate = False
'        Else
'            Panel.Bevel = sbrInset
'            Panel.Style = sbrDate
'            Panel.ToolTipText = "Click to hide date"
'            showDate = True
'        End If
    Else
    End If

End Sub

Sub remSettings()

    On Error Resume Next

    If delSet = False Then
        currHt = Me.Height
        currWd = Me.Width
        currTop = Me.Top
        currLft = Me.Left
        setRegSettings
    Else
        deleteAllSettings
    End If

End Sub

Sub setStatBar()

On Error Resume Next

    If TopMost = True Then
        StatusBar1.Panels(1).Text = "Always on top"
        StatusBar1.Panels(1).Bevel = sbrRaised
        StatusBar1.Panels(1).ToolTipText = "Click to make window dockable"
        Unload frmOnTop
    ElseIf TopMost = False Then
        StatusBar1.Panels(1).Text = "Dockable"
        StatusBar1.Panels(1).Bevel = sbrInset
        StatusBar1.Panels(1).ToolTipText = "Click to be always on top"
        Load frmOnTop
        SetFrmOnTop
        frmOnTop.Show
    End If
    
    If showTime = True Then
        StatusBar1.Panels(2).Style = sbrTime
        StatusBar1.Panels(2).Bevel = sbrInset
        StatusBar1.Panels(2).ToolTipText = "Click to hide time"
    ElseIf showTime = False Then
        StatusBar1.Panels(2).Style = sbrText
        StatusBar1.Panels(2).Text = ""
        StatusBar1.Panels(2).Bevel = sbrNoBevel
        StatusBar1.Panels(2).ToolTipText = "Click to show time"
    End If
    
    If showDate = True Then
        StatusBar1.Panels(3).Style = sbrDate
        StatusBar1.Panels(3).Bevel = sbrInset
        StatusBar1.Panels(3).ToolTipText = "Click to hide date"
    ElseIf showTime = False Then
        StatusBar1.Panels(3).Style = sbrText
        StatusBar1.Panels(3).Text = ""
        StatusBar1.Panels(3).Bevel = sbrNoBevel
        StatusBar1.Panels(3).ToolTipText = "Click to show date"
    End If

End Sub

Sub comeBackWhere()

    On Error Resume Next

    Select Case comeBack
    Case 1:
        mnuOpenItem_Click
    Case 2:
        mnuNewItem_Click
    End Select

End Sub

Sub wantToSave()

        On Error Resume Next

        Load frmMsgBoxYNC
        frmMsgBoxYNC.Caption = "TextPad"
        frmMsgBoxYNC.lblTop = "The text in the " & saveString & " has changed."
        frmMsgBoxYNC.lblBottom = "Do you want to save the changes?"
        frmMsgBoxYNC.Show vbModal
        qt = quitConfirm
        If qt = vbYes Then
            If firstSave = True Then
                save_It_As
            Else
                Save_It
            End If
            savedIt = True
            setFileInfo (10)
        ElseIf qt = vbNo Then
            savedIt = True
        ElseIf qt = vbCancel Then
            Exit Sub
        End If
         
        comeBackWhere

End Sub

Private Sub save_It_As()

    
    On Error Resume Next
    
    cmnDlg1.FileName = saveString
    'cmnDlg1.CancelError = True
    cmnDlg1.Filter = "Text Document (*.txt)|*.txt|Rich Text Document (*.rtf) |*.rtf|All Files|*.*"
    cmnDlg1.ShowSave
    saveFilePath = cmnDlg1.FileName
    saveString = cmnDlg1.FileTitle
    
    If saveString <> "" Then
        rtbox1.SaveFile saveString, rtfText
        firstSave = False
        savedIt = True
        beingOpened = True
        setFileInfo (10)
    End If
    
'errHand:
'    If Err.Number = 32755 Then
'
'        Exit Sub
'    Else
'        Resume Next
'    End If

End Sub

Private Sub Save_It()

        On Error Resume Next
    
        rtbox1.SaveFile saveString
        firstSave = False
        savedIt = True
        beingOpened = True
        setFileInfo (10)

End Sub

Sub setFileInfo(theCase)

    On Error Resume Next

    Select Case theCase
    Case 0:
        StatusBar1.Panels(6).Text = "Untitled"
    Case 10:
        If savedIt = True Then
            StatusBar1.Panels(6).Text = "Saved"
        Else
            StatusBar1.Panels(6).Text = "Unsaved"
        End If
    
    End Select

End Sub

Sub setLineNum()

    Dim currLine As Long

    On Local Error Resume Next
    currLine = SendMessage(rtbox1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    StatusBar1.Panels(7).Text = "Line: " + Format$(currLine, "#,###,###")


End Sub

Sub uncheckViewSize()

    mnuXSmallItem.Checked = False
    mnuSmallItem.Checked = False
    mnuMediumItem.Checked = False
    mnuLargeItem.Checked = False
    mnuXLargeItem.Checked = False

End Sub

''''''''''''Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''''''''''''
''''''''''''Select Case Button.Key
''''''''''''Case "bold"
''''''''''''    mnuBoldAllItem_Click
''''''''''''    Button.Image = 1
''''''''''''    Button.Key = "bold_down"
''''''''''''Case "bold_down"
''''''''''''    Button.Image = 2
''''''''''''    Button.Key = "bold"
''''''''''''    'unbold
''''''''''''Case "italic"
''''''''''''    mnuItalicAllItem_Click
''''''''''''    Button.Image = 3
''''''''''''    Button.Key = "italic_down"
''''''''''''Case "italic_down"
''''''''''''    'unital
''''''''''''    Button.Image = 4
''''''''''''    Button.Key = "italic"
''''''''''''Case "uline"
''''''''''''    mnuULAllItem_Click
''''''''''''    Button.Image = 5
''''''''''''    Button.Key = "uline_down"
''''''''''''Case "uline_down"
''''''''''''    Button.Image = 6
''''''''''''    Button.Key = "uline"
''''''''''''End Select
''''''''''''
''''''''''''End Sub
''''''''''''

Sub setPause(pauseTime)

   start = Timer
   
   Do While Timer < start + pauseTime
      DoEvents
   Loop

End Sub


Private Sub tbr1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1:
        mnuOpenItem_Click
    Case 2:
        mnuNewItem_Click
    Case 3:
        mnuSaveItem_Click
    Case 4:
        mnuPrintItem_Click
    Case 6:
        mnuChkSpellItem_Click
    Case 8:
        mnuCopyItem_Click
    Case 9:
        mnuCutItem_Click
    Case 10:
        mnuPasteItem_Click
    Case 11:
        mnuHdnDelItem_Click
    Case 13:
        mnuCBViewItem_Click
    Case 15:
        mnuUndoItem_Click
    Case 16:
        mnuRedoItem_Click
    Case 18:
        mnuBoldItem_Click
    Case 19:
        mnuItalicItem_Click
    Case 20:
        mnuULItem_Click
    Case 21:
        mnuSTItem_Click
    Case 23:
        mnuFindItem_Click
    Case 25:
        rtbox1.SelStart = 0
        rtbox1.SelLength = Len(rtbox1.Text)
        uc = UCase(rtbox1.SelText)
        rtbox1.Text = uc
        
        rtbox1.SelStart = Len(rtbox1.Text)
        rtbox1.SelLength = Len(rtbox1.Text)
        rtbox1.SelLength = 0
    Case 26:
        rtbox1.SelStart = 0
        rtbox1.SelLength = Len(rtbox1.Text)
        lc = LCase(rtbox1.SelText)
        rtbox1.Text = lc
        
        rtbox1.SelStart = Len(rtbox1.Text)
        rtbox1.SelLength = Len(rtbox1.Text)
        rtbox1.SelLength = 0
    End Select

End Sub

Private Sub tbr1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    If ButtonMenu.Key = "tbrNewItm" Then
        mnuNewPadItem_Click
    ElseIf ButtonMenu.Key = "tbrSameItm" Then
        mnuNewItem_Click
    ElseIf ButtonMenu.Key = "ucaseAll" Then
        rtbox1.SelStart = 0
        rtbox1.SelLength = Len(rtbox1.Text)
        uc = UCase(rtbox1.SelText)
        rtbox1.Text = uc
        
        rtbox1.SelStart = Len(rtbox1.Text)
        rtbox1.SelLength = Len(rtbox1.Text)
        rtbox1.SelLength = 0
    ElseIf ButtonMenu.Key = "ucaseSeled" Then
        uc = UCase(rtbox1.SelText)
        rtbox1.SelText = uc
    ElseIf ButtonMenu.Key = "lcaseAll" Then
        rtbox1.SelStart = 0
        rtbox1.SelLength = Len(rtbox1.Text)
        LCase (rtbox1.SelText)
        lc = LCase(rtbox1.SelText)
        rtbox1.Text = lc
        
        rtbox1.SelStart = Len(rtbox1.Text)
        rtbox1.SelLength = Len(rtbox1.Text)
        rtbox1.SelLength = 0
    ElseIf ButtonMenu.Key = "lcaseSeled" Then
        lc = LCase(rtbox1.SelText)
        rtbox1.SelText = lc
    End If

End Sub

Private Sub tbr2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "calendar"
        mnuShowCalItem_Click
    Case "timer"
        mnuSetTmrItem_Click
    Case "calc"
        mnuCalcItem_Click
    Case "lock"
        mnuLockItem_Click
    Case "unlock"
        mnuLockItem_Click
    Case "options"
        mnuSettingItem_Click
    Case "font"
        mnuFontNameItem_Click
    Case "about"
        mnuAboutItem_Click
    Case "help"
        mnuAllHelpItem_Click
    End Select

End Sub

Private Sub tpSearch_Change()

    If tpSearch.Text <> "" Then
        tpReplace.BackColor = &HC0FFFF
        tpReplace.Enabled = True
    Else
        tpReplace.BackColor = &HC0C0C0
        tpReplace.Enabled = False
    End If

End Sub
