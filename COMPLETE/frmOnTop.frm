VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOnTop 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   1140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   345
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOnTop.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOnTop.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOnTop.frx":02B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   870
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   60
         TabIndex        =   1
         Top             =   105
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "textPad"
               Object.ToolTipText     =   "Bring TextPad on top"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "close"
               Object.ToolTipText     =   "Close Me"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   765
      Left            =   15
      Top             =   15
      Width           =   1110
   End
End
Attribute VB_Name = "frmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Me.Top = 235
    Me.Left = Screen.Width - Me.Width
    
    onTopActive = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

     onTopActive = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "close" Then
        Unload Me
    Else
        If frmTextPad.WindowState = vbMinimized Then
            frmTextPad.WindowState = vbNormal
            frmTextPad.Height = currHt
            frmTextPad.Width = currWd
            frmTextPad.Top = currTop
            frmTextPad.Left = currLft
        End If
        frmTextPad.rtbox1.SetFocus
    End If

End Sub
