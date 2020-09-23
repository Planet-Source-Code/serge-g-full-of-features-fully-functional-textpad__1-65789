VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSpellCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Check"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "spellCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Your own"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   5010
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdSpell 
      Caption         =   "Spell Check"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"spellCheck.frx":014A
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "If not listed"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As New Word.Application
Dim wdsp As Word.SpellingSuggestions
Dim z As Long
Dim errs As Long
Dim cancelSave As Boolean

Private Sub cmdSpell_Click()

    Dim theT As Variant
    Dim temp As String
    Dim newTemp As String
    
    newTemp = " " & rtb1.Text
    newTemp = Replace(newTemp, Chr(13) & Chr(10), " ", , , vbTextCompare)
'    temp = Replace(temp, ",", " ", , , vbTextCompare)
'    temp = Replace(temp, ".", " ", , , vbTextCompare)
'    temp = Replace(temp, "!", " ", , , vbTextCompare)
'    temp = Replace(temp, "?", " ", , , vbTextCompare)
    newTemp = removeDelimiters(newTemp, ",", ".", "?", "!")
    temp = removeSpaces(newTemp)
    theT = Split(temp, " ")
    
    If z > UBound(theT) Then
        cmdSpell.Caption = "Spell Check"
        a = MsgBox("Spell check complete, start again?", vbYesNo, "Spell Check")
        If a = vbNo Then
            Me.Height = 4100
            List1.Visible = False
            Exit Sub
        Else
            z = 1
        End If
    End If
    
    cmdSpell.Enabled = False
    
    Dim R As Integer
    R = rtb1.Find(theT(z), 0, , rtfWholeWord)
    If R <> -1 Then
        rtb1.SelStart = R
        rtb1.SelLength = Len(theT(z))
        rtb1.SelColor = vbRed
        If Not X.CheckSpelling(theT(z)) Then
            List1.Visible = True
            Me.Height = 5880
            errs = errs + 1
            X.Documents.Add
            Set wdsp = X.GetSpellingSuggestions(theT(z))
            List1.Clear
            For p = 1 To wdsp.Count
                List1.AddItem wdsp(p).Name
                List1.Tag = theT(z)
            Next p
            cmdSpell.Caption = "Check Next"
            cmdSpell.Enabled = True
        Else
            rtb1.SelColor = vbBlack
            If z >= UBound(theT) Then
                cmdSpell.Enabled = True
                cmdSpell.Caption = "Spell Check"
                rtb1.SelStart = 0
                rtb1.SelLength = Len(rtb1.Text)
                rtb1.SelColor = vbBlack
                If errs = 0 Then
                    MsgBox ("Spell check complete, no errors")
                Else
                    MsgBox ("Spell check complete, total errors found : " & errs)
                    errs = 0
                End If
                Me.Height = 4100
                List1.Visible = False
                Exit Sub
            Else
                z = z + 1
                cmdSpell_Click
            End If
        End If
    End If
    
'    If z >= UBound(theT) + 1 Then
'        MsgBox ("Spell check complete")
'        cmdSpell.Caption = "Spell Check"
'        cmdSpell.Enabled = True
'        Exit Sub
'    Else
'        z = z + 1
'    End If
    
End Sub

Private Sub Command1_Click()

    If rtb1.SelLength > 0 Then
        rtb1.SelText = Text1.Text
        Text1.Text = ""
    End If

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Command3_Click()

    cancelSave = True
    Unload Me

End Sub

Private Sub Form_Load()
    
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    
    cancelSave = False
    
    tempStat = TopMost
    TopMost = False
    SetTopMost

    rtb1.Text = frmTextPad.rtbox1.Text

    z = 1
    errs = 0
    Me.Height = 4100
    List1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If cancelSave = False Then
        frmTextPad.rtbox1.Text = rtb1.Text
    End If
    
    TopMost = tempStat
    SetTopMost

End Sub

Private Sub List1_Click()

    f = rtb1.Find(List1.Tag, , , rtfWholeWord)
    If f <> -1 Then
        rtb1.SelStart = f
        rtb1.SelLength = Len(List1.Tag)
        rtb1.SelColor = vbBlack
        rtb1.SelText = List1.Text
        List1.Tag = List1.Text
    Else
    End If

End Sub

Private Sub rtb1_KeyPress(KeyAscii As Integer)

    z = 1
    errs = 0
    cmdSpell.Enabled = True
    cmdSpell.Caption = "Spell Check"

End Sub

Function removeDelimiters(theText As String, ParamArray delim()) As String

    Dim temp As String
    Dim del As Variant
    
    temp = theText
    
    For Each del In delim
        temp = Replace(temp, del, " ")
    Next del
    
    removeDelimiters = temp

End Function

Function removeSpaces(theText As String) As String

    Dim temp
    Dim flg As Integer
    flg = 0
    
    temp = theText
    
    Do Until InStr(temp, "  ") = 0
        flg = flg + 1
        If flg >= 250 Then End
        temp = Replace(temp, "  ", " ")
    Loop
    
    temp = LTrim(temp)
    
    removeSpaces = temp

End Function

