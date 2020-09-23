VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1035
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Changed Caption"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Host's Caption"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objHost As Object

Private Sub Command3_Click()
objHost.Caption = Text1
End Sub
