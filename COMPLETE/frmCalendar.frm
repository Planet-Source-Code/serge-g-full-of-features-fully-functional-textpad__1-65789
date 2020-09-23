VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendar 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4875
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   8370
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Insert Date"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Time"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3390
      Left            =   270
      TabIndex        =   1
      Top             =   510
      Width           =   3645
      _Version        =   524288
      _ExtentX        =   6429
      _ExtentY        =   5980
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2006
      Month           =   6
      Day             =   19
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   2
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calendar"
      Height          =   3810
      Left            =   150
      TabIndex        =   0
      Top             =   255
      Width           =   3900
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5370
      Top             =   5865
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6135
      Top             =   5865
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      Height          =   270
      Left            =   4515
      TabIndex        =   2
      Top             =   255
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   3735
      Left            =   4320
      Top             =   345
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   3735
      Left            =   4335
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   16
      Left            =   5115
      ToolTipText     =   "Move"
      Top             =   2085
      Width           =   225
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2145
      Width           =   135
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   5115
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   18
      Left            =   4425
      MouseIcon       =   "frmCalendar.frx":030A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Help"
      Top             =   5730
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   19
      Left            =   4155
      MouseIcon       =   "frmCalendar.frx":045C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Hide For 10 Seconds"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   20
      Left            =   4665
      MouseIcon       =   "frmCalendar.frx":05AE
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   5910
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   14
      Left            =   4155
      MouseIcon       =   "frmCalendar.frx":0700
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   15
      Left            =   4665
      MouseIcon       =   "frmCalendar.frx":0852
      MousePointer    =   99  'Custom
      Top             =   6180
      Width           =   150
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   152
      X2              =   234
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   43
      X2              =   125
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   149
      X2              =   234
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   41
      X2              =   126
      Y1              =   409
      Y2              =   409
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Index           =   1
      X1              =   151
      X2              =   233
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      Index           =   0
      X1              =   43
      X2              =   125
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   59
      Left            =   6465
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   58
      Left            =   6330
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   57
      Left            =   6165
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   56
      Left            =   6015
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   55
      Left            =   5865
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   54
      Left            =   5715
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   53
      Left            =   5550
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   52
      Left            =   5415
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   51
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   50
      Left            =   5145
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   49
      Left            =   5010
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   48
      Left            =   4860
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   47
      Left            =   4710
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   46
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   45
      Left            =   4410
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   44
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   43
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   42
      Left            =   3945
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   41
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   40
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   39
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   38
      Left            =   3315
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   37
      Left            =   3135
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   36
      Left            =   2925
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   35
      Left            =   2745
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   34
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   33
      Left            =   2415
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   32
      Left            =   2220
      Shape           =   3  'Circle
      Top             =   5175
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   31
      Left            =   2025
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   30
      Left            =   1845
      Shape           =   3  'Circle
      Top             =   5205
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   29
      Left            =   1665
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   28
      Left            =   1485
      Shape           =   3  'Circle
      Top             =   5235
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   27
      Left            =   1335
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   26
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   5250
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   25
      Left            =   975
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   24
      Left            =   810
      Shape           =   3  'Circle
      Top             =   5265
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   23
      Left            =   660
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   22
      Left            =   6450
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6255
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   20
      Left            =   6060
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   5895
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   5730
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   5340
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   5190
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   15
      Left            =   5055
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   4875
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   4665
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   4425
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   4275
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   10
      Left            =   4095
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   3930
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   5460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   3540
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   3255
      Shape           =   3  'Circle
      Top             =   5475
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   5
      Left            =   3015
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   2835
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   2325
      Shape           =   3  'Circle
      Top             =   5505
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   2100
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   150
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempStat As Boolean
Const PI = 3.141592654

Private Sub Calendar1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Command2_Click()

    frmTextPad.mnuInTimeItem_Click

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Command3_Click()

    frmTextPad.mnuInDateItem_Click

End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
  
  tempStat = TopMost
  TopMost = False
  SetTopMost
  
  Dim Ret As Long
  Dim CLR As Long
  Me.Hide
  
  Me.Top = Screen.Height / 2 - (Me.Height / 2)
  Me.Left = Screen.Width / 2 - (Me.Width / 2)
  
  Shape3.Left = 411
  Shape3.Top = 140
  
  Shape2.Left = 414
  Shape2.Top = 143
  
  For i = 0 To 59
    If i Mod 5 = 0 Then
      Shape4(i).Left = 418 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
      Shape4(i).Top = 148 + Sin(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
    Else
      Shape4(i).Left = 418 + Cos(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
      Shape4(i).Top = 148 + Sin(i * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
    End If
    Shape4(i).BorderColor = QBColor(15)
    Shape4(i).FillColor = &H808080         '&H8000000A
  Next i
  Image2(18).Left = 418 + Cos(0 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(18).Top = 148 + Sin(0 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(19).Left = 418 + Cos(50 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(19).Top = 148 + Sin(50 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(20).Left = 418 + Cos(10 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(20).Top = 148 + Sin(10 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(14).Left = 418 + Cos(40 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(14).Top = 148 + Sin(40 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(15).Left = 418 + Cos(20 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  Image2(15).Top = 148 + Sin(20 * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
  
  Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

    TopMost = tempStat
    SetTopMost

End Sub

Private Sub Timer1_Timer()
  Dim Tim As Long
  Tim = Int(Timer)
  For i = 0 To 1
    'Hour
    Line1(i).X1 = 418
    Line1(i).Y1 = 148
    Line1(i).X2 = 418 + Cos((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
    Line1(i).Y2 = 148 + Sin((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
    'Minute
    Line2(i).X1 = 418
    Line2(i).Y1 = 148
    Line2(i).X2 = 418 + Cos((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    Line2(i).Y2 = 148 + Sin((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    'Second
    Line3(i).X1 = 418 - Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
    Line3(i).Y1 = 148 - Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
    Line3(i).X2 = 418 + Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    Line3(i).Y2 = 148 + Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
  Next i
End Sub
Private Sub Timer2_Timer()
  DelayCounter = DelayCounter + 1
  If DelayCounter >= 100 Then
     frmCalendar.Show
     Timer2.Interval = 0
  End If
End Sub
