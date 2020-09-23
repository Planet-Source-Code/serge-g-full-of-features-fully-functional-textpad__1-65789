Attribute VB_Name = "Module1"
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SendMessage Lib _
    "User32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long


Public TopMost As Boolean
Public currWd As Single
Public currHt As Single
Public currTop As Single
Public currLft As Single
Public currViewSize As Integer
Public showTime As Boolean
Public showDate As Boolean
Public openFileCount As Integer
Public exeLocation As String
Public quitConfirm As VbMsgBoxResult
Public specialCopyChar As String
Public iAmActive As Integer
Public topMostTemp As Boolean
Public publicCopy  As Variant
Public currFontName As String
Public currFontSize As Single
Public bullIndentBy As Integer
Public indentLeft As Integer
Public indentRight As Integer
Public defaultFont As String
Public defaultBGColor As Long
Public defaultFontColor As Long
Public currViewName As String
Public defaultIndentAll As String
Public customWordCount As Integer
Public customWord(1 To 5) As String
Public isQFActive As Boolean
Public useLastChar As Boolean
Public isTimerSet As Boolean
Private dataExists As Boolean
Public dataFileName As Variant
Public onTopActive As Boolean

Public Const EM_LINEFROMCHAR = &HC9

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


Public Sub SetTopMost()

    On Error Resume Next
  
  If onTopActive = False Then
  
    If TopMost Then
        SetWindowPos frmTextPad.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        frmTextPad.StatusBar1.Panels(1).Bevel = sbrRaised
        frmTextPad.StatusBar1.Panels(1).Text = "Always on top"
        frmTextPad.StatusBar1.Panels(1).ToolTipText = "Click to make window dockable"
    Else
        SetWindowPos frmTextPad.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        frmTextPad.StatusBar1.Panels(1).Bevel = sbrInset
        frmTextPad.StatusBar1.Panels(1).Text = "Dockable"
        frmTextPad.StatusBar1.Panels(1).ToolTipText = "Click to be always on top"
    End If
    
  End If
    
End Sub

Public Sub SetFrmOnTop()

    On Error Resume Next

        SetWindowPos frmOnTop.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        
End Sub


Public Sub searchOnTop()

    On Error Resume Next

    topMostTemp = TopMost
    
    If TopMost Then
        TopMost = False
        SetTopMost
    End If
    
    SetWindowPos frmSearchTxt.hWnd, -1, 0, 0, 0, 0, 1 Or 2

End Sub

Public Sub searchNotTop()

    SetWindowPos frmSearchTxt.hWnd, -2, 0, 0, 0, 0, 1 Or 2

End Sub


Public Sub setRegSettings()

    On Error Resume Next
    
    SaveSetting appname:="TextPad", section:="Startup", Key:="Window_Width", setting:=currWd
    SaveSetting appname:="TextPad", section:="Startup", Key:="Window_Height", setting:=currHt
    SaveSetting appname:="TextPad", section:="Startup", Key:="Window_Top", setting:=currTop
    SaveSetting appname:="TextPad", section:="Startup", Key:="Window_Left", setting:=currLft
    SaveSetting appname:="TextPad", section:="Startup", Key:="Window_OnTop", setting:=TopMost
    SaveSetting appname:="TextPad", section:="Startup", Key:="StatusBar_ShowTime", setting:=showTime
    SaveSetting appname:="TextPad", section:="Startup", Key:="StatusBar_ShowDate", setting:=showDate
    SaveSetting appname:="TextPad", section:="Startup", Key:="View_Size", setting:=currViewSize
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_Font", setting:=defaultFont
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_Font_Color", setting:=defaultFontColor
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_BG_Color", setting:=defaultBGColor
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Bullet_Indent_By", setting:=bullIndentBy
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Left_Indent_By", setting:=indentLeft
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Right_Indent_By", setting:=indentRight
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="isDefault_Indent", setting:=defaultIndentAll
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Auto_Uppercase", setting:=useLastChar
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Dat_Exists", setting:=dataExists
    SaveSetting appname:="TextPad", section:="User_Settings2", Key:="Dat_File_Name", setting:=dataFileName

End Sub

Public Sub getRegSettings()

    On Error Resume Next

    currWd = GetSetting(appname:="TextPad", section:="Startup", Key:="Window_Width", Default:="6690")
    currHt = GetSetting(appname:="TextPad", section:="Startup", Key:="Window_Height", Default:="5565")
    currTop = GetSetting(appname:="TextPad", section:="Startup", Key:="Window_Top", Default:="-10")
    currLft = GetSetting(appname:="TextPad", section:="Startup", Key:="Window_Left", Default:="-10")
    TopMost = CBool(GetSetting(appname:="TextPad", section:="Startup", Key:="Window_OnTop", Default:="False"))
    showTime = CBool(GetSetting(appname:="TextPad", section:="Startup", Key:="StatusBar_ShowTime", Default:="True"))
    showDate = CBool(GetSetting(appname:="TextPad", section:="Startup", Key:="StatusBar_ShowDate", Default:="True"))
    currViewSize = CInt(GetSetting(appname:="TextPad", section:="Startup", Key:="View_Size", Default:="9"))
    defaultFont = GetSetting(appname:="TextPad", section:="Startup", Key:="Default_Font", Default:="MS Sans Serif")
    defaultFontColor = CLng(GetSetting(appname:="TextPad", section:="Startup", Key:="Default_Font_Color", Default:="986895"))
    defaultBGColor = CLng(GetSetting(appname:="TextPad", section:="Startup", Key:="Default_BG_Color", Default:="16777215"))
    bullIndentBy = CInt(GetSetting(appname:="TextPad", section:="User_Settings", Key:="Bullet_Indent_By", Default:="500"))
    indentLeft = CInt(GetSetting(appname:="TextPad", section:="User_Settings", Key:="Left_Indent_By", Default:="500"))
    indentRight = CInt(GetSetting(appname:="TextPad", section:="User_Settings", Key:="Right_Indent_By", Default:="500"))
    defaultIndentAll = GetSetting(appname:="TextPad", section:="User_Settings", Key:="isDefault_Indent", Default:="No Indent")
    useLastChar = CBool(GetSetting(appname:="TextPad", section:="User_Settings", Key:="Auto_Uppercase", Default:="False"))
    dataExists = CBool(GetSetting(appname:="TextPad", section:="User_Settings", Key:="Dat_Exists", Default:="False"))
    dataFileName = GetSetting(appname:="TextPad", section:="User_Settings2", Key:="Dat_File_Name", Default:="CustomText.dat")

    getData

End Sub

Sub deleteAllSettings()
    
    On Error Resume Next
    
    DeleteSetting appname:="TextPad", section:="Startup"
    DeleteSetting appname:="TextPad", section:="User_Settings"
    
    Open dataFileName For Output As #1
    Close #1

End Sub

Public Sub anyOpenFiles()

    openFileCount = GetSetting(appname:="TextPad", section:="Startup", Key:="AnyOpenFiles", Default:="1")
    
End Sub

Public Sub writeTempReg()
    
    On Error Resume Next
    
    SaveSetting appname:="TextPad", section:="Startup", Key:="AnyOpenFiles", setting:=openFileCount + 1

End Sub

Public Sub deleteTempReg()

    On Error Resume Next
    
    DeleteSetting appname:="TextPad", section:="Startup", Key:="AnyOpenFiles"
        
End Sub

Public Sub saveFileLocation()
    
    On Error Resume Next
    
    exeLocation = App.Path & "\" & App.EXEName
    SaveSetting appname:="TextPad", section:="Startup", Key:="File_Location", setting:=exeLocation
            
End Sub

Public Sub mySavedOptions()

    On Error Resume Next

    specialCopyChar = " "     'Change later
    
    Select Case currViewSize
    Case 8:
        frmTextPad.mnuXSmallItem.Checked = True
        currViewName = "Extra Small"
    Case 9:
        frmTextPad.mnuSmallItem.Checked = True
        currViewName = "Small"
    Case 11:
        frmTextPad.mnuMediumItem.Checked = True
        currViewName = "Medium"
    Case 13:
        frmTextPad.mnuLargeItem.Checked = True
        currViewName = "Large"
    Case 16:
        frmTextPad.mnuXLargeItem.Checked = True
        currViewName = "Extra Large"
    Case Else:
        currViewSize = 10
    End Select
    
    frmTextPad.rtbox1.Font.Size = currViewSize

End Sub

Public Sub saveOptions()

    On Error Resume Next

    SaveSetting appname:="TextPad", section:="Startup", Key:="StatusBar_ShowTime", setting:=showTime
    SaveSetting appname:="TextPad", section:="Startup", Key:="StatusBar_ShowDate", setting:=showDate
    SaveSetting appname:="TextPad", section:="Startup", Key:="View_Size", setting:=currViewSize
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_Font", setting:=defaultFont
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_Font_Color", setting:=defaultFontColor
    SaveSetting appname:="TextPad", section:="Startup", Key:="Default_BG_Color", setting:=defaultBGColor
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Bullet_Indent_By", setting:=bullIndentBy
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Left_Indent_By", setting:=indentLeft
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Right_Indent_By", setting:=indentRight
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="isDefault_Indent", setting:=defaultIndentAll
    SaveSetting appname:="TextPad", section:="User_Settings", Key:="Auto_Uppercase", setting:=useLastChar

End Sub

Public Sub customWordMenu()

    On Error Resume Next

    setCaption

    If Len(customWord(1)) > 10 Then
        frmTextPad.mnuWord1Item.Caption = Mid(customWord(1), 1, 10) & "..."
    Else
        frmTextPad.mnuWord1Item.Caption = customWord(1)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(customWord(2)) > 10 Then
        frmTextPad.mnuWord2Item.Caption = Mid(customWord(2), 1, 10) & "..."
    Else
        frmTextPad.mnuWord2Item.Caption = customWord(2)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(customWord(3)) > 10 Then
        frmTextPad.mnuWord3Item.Caption = Mid(customWord(3), 1, 10) & "..."
    Else
        frmTextPad.mnuWord3Item.Caption = customWord(3)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(customWord(4)) > 10 Then
        frmTextPad.mnuWord4Item.Caption = Mid(customWord(4), 1, 10) & "..."
    Else
        frmTextPad.mnuWord4Item.Caption = customWord(4)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(customWord(5)) > 10 Then
        frmTextPad.mnuWord1Item.Caption = Mid(customWord(5), 1, 10) & "..."
    Else
        frmTextPad.mnuWord5Item.Caption = customWord(5)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Select Case customWordCount
    Case 1:
        frmTextPad.mnuCustomInsert.Enabled = True
        frmTextPad.mnuWord1Item.Enabled = True
        frmTextPad.mnuWord2Item.Enabled = False
          frmTextPad.mnuWord2Item.Caption = "Word 2"
        frmTextPad.mnuWord3Item.Enabled = False
          frmTextPad.mnuWord3Item.Caption = "Word 3"
        frmTextPad.mnuWord4Item.Enabled = False
          frmTextPad.mnuWord4Item.Caption = "Word 4"
        frmTextPad.mnuWord5Item.Enabled = False
          frmTextPad.mnuWord5Item.Caption = "Word 5"
    Case 2:
        frmTextPad.mnuCustomInsert.Enabled = True
        frmTextPad.mnuWord1Item.Enabled = True
        frmTextPad.mnuWord2Item.Enabled = True
        frmTextPad.mnuWord3Item.Enabled = False
          frmTextPad.mnuWord3Item.Caption = "Word 3"
        frmTextPad.mnuWord4Item.Enabled = False
          frmTextPad.mnuWord4Item.Caption = "Word 4"
        frmTextPad.mnuWord5Item.Enabled = False
          frmTextPad.mnuWord5Item.Caption = "Word 5"
    Case 3:
        frmTextPad.mnuCustomInsert.Enabled = True
        frmTextPad.mnuWord1Item.Enabled = True
        frmTextPad.mnuWord2Item.Enabled = True
        frmTextPad.mnuWord3Item.Enabled = True
        frmTextPad.mnuWord4Item.Enabled = False
          frmTextPad.mnuWord4Item.Caption = "Word 4"
        frmTextPad.mnuWord5Item.Enabled = False
          frmTextPad.mnuWord5Item.Caption = "Word 5"
    Case 4:
        frmTextPad.mnuCustomInsert.Enabled = True
        frmTextPad.mnuWord1Item.Enabled = True
        frmTextPad.mnuWord2Item.Enabled = True
        frmTextPad.mnuWord3Item.Enabled = True
        frmTextPad.mnuWord4Item.Enabled = True
        frmTextPad.mnuWord5Item.Enabled = False
          frmTextPad.mnuWord5Item.Caption = "Word 5"
    Case 5:
        frmTextPad.mnuCustomInsert.Enabled = True
        frmTextPad.mnuWord1Item.Enabled = True
        frmTextPad.mnuWord2Item.Enabled = True
        frmTextPad.mnuWord3Item.Enabled = True
        frmTextPad.mnuWord4Item.Enabled = True
        frmTextPad.mnuWord5Item.Enabled = True
    Case Else
        frmTextPad.mnuCustomInsert.Enabled = False
        frmTextPad.mnuCustomInsert.Enabled = False
        frmTextPad.mnuWord1Item.Enabled = False
        frmTextPad.mnuWord2Item.Enabled = False
        frmTextPad.mnuWord3Item.Enabled = False
        frmTextPad.mnuWord4Item.Enabled = False
        frmTextPad.mnuWord5Item.Enabled = False
    End Select

    saveToDataFile

End Sub

Private Sub setCaption()

    frmTextPad.mnuWord1Item.Caption = "Word 1"
    frmTextPad.mnuWord2Item.Caption = "Word 2"
    frmTextPad.mnuWord3Item.Caption = "Word 3"
    frmTextPad.mnuWord4Item.Caption = "Word 4"
    frmTextPad.mnuWord5Item.Caption = "Word 5"

End Sub

Public Sub saveToDataFile()

    On Error Resume Next

    Open dataFileName For Output As #1
        For z = 1 To customWordCount
            Print #1, customWord(z)
            dataExists = True
        Next z
    Close #1

End Sub

Public Sub getData()

On Error GoTo erHand

    X = 0
    
    Open dataFileName For Input As #1
    Do While Not EOF(1)
        X = X + 1
        customWordCount = X
        Input #1, customWord(X)
    Loop
    Close #1

erHand:
If Err.Number = 53 And dataExists = False Then
    Open dataFileName For Output As #1
    Close #1
    Resume Next
ElseIf Err.Number = 53 And dataExists = True Then
    findIt = MsgBox("The file containing your custom text was moved or deleted. Do you want to find it?", vbYesNo, "Data Not Found")
        If findIt = vbNo Then
            Open dataFileName For Output As #1
            Close #1
            Resume Next
        ElseIf findIt = vbYes Then
            frmTextPad.cmndlg2.Filter = "Data File (*.dat)|*.dat|)"
            frmTextPad.cmndlg2.DialogTitle = "Locate the .dat file"
            frmTextPad.cmndlg2.ShowOpen
                
                If frmTextPad.cmndlg2.FileName <> "" Then
                    dataFileName = frmTextPad.cmndlg2.FileName
                    Open dataFileName For Input As #1
                        Do While Not EOF(1)
                            X = X + 1
                            customWordCount = X
                            Input #1, customWord(X)
                        Loop
                    Close #1

                Else
                    '''New file will be created
                End If
            Resume Next
        End If
End If
    

End Sub
