VERSION 5.00
Begin VB.Form frmTextPad_ImageList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ImageList"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1710
   Icon            =   "frmTextPad_ImageList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image img_Underline_Mixed 
      Height          =   270
      Left            =   1080
      Picture         =   "frmTextPad_ImageList.frx":0442
      Top             =   1440
      Width           =   270
   End
   Begin VB.Image img_Strike_Mixed 
      Height          =   270
      Left            =   1080
      Picture         =   "frmTextPad_ImageList.frx":0874
      Top             =   1080
      Width           =   270
   End
   Begin VB.Image img_Italic_Mixed 
      Height          =   270
      Left            =   1080
      Picture         =   "frmTextPad_ImageList.frx":0CA6
      Top             =   720
      Width           =   270
   End
   Begin VB.Image img_Bold_Mixed 
      Height          =   270
      Left            =   1080
      Picture         =   "frmTextPad_ImageList.frx":10D8
      Top             =   360
      Width           =   270
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1815
      Left            =   135
      Top             =   135
      Width           =   1440
   End
   Begin VB.Image img_Strike_Raised 
      Height          =   270
      Left            =   720
      Picture         =   "frmTextPad_ImageList.frx":150A
      Top             =   1080
      Width           =   270
   End
   Begin VB.Image img_Strike_Inset 
      Height          =   270
      Left            =   360
      Picture         =   "frmTextPad_ImageList.frx":193C
      Top             =   1080
      Width           =   270
   End
   Begin VB.Image img_Italic_Raised 
      Height          =   270
      Left            =   720
      Picture         =   "frmTextPad_ImageList.frx":1D6E
      Top             =   720
      Width           =   270
   End
   Begin VB.Image img_Italic_Inset 
      Height          =   270
      Left            =   360
      Picture         =   "frmTextPad_ImageList.frx":21A0
      Top             =   720
      Width           =   270
   End
   Begin VB.Image img_Bold_Raised 
      Height          =   270
      Left            =   720
      Picture         =   "frmTextPad_ImageList.frx":25D2
      Top             =   360
      Width           =   270
   End
   Begin VB.Image img_Bold_Inset 
      Height          =   270
      Left            =   360
      Picture         =   "frmTextPad_ImageList.frx":2A04
      Top             =   360
      Width           =   270
   End
   Begin VB.Image img_Underline_Inset 
      Height          =   270
      Left            =   360
      Picture         =   "frmTextPad_ImageList.frx":2E36
      Top             =   1440
      Width           =   270
   End
   Begin VB.Image img_Underline_Raised 
      Height          =   270
      Left            =   720
      Picture         =   "frmTextPad_ImageList.frx":3268
      Top             =   1440
      Width           =   270
   End
End
Attribute VB_Name = "frmTextPad_ImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

