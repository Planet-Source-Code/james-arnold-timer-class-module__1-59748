VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timer"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Elapsed"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTimer As timTimer 'Our timer

Private Sub Command1_Click()

  Command1.Enabled = False
  Command2.Enabled = True
  Command3.Enabled = True
  
  Set cTimer = New timTimer 'Assign a new timer
  cTimer.StartTimer 'Start it
  
End Sub

Private Sub Command2_Click()

  'Display elapsed time
  MsgBox FormatNumber(cTimer.Elapsed, 0) & " milliseconds have passed, or " & _
  FormatNumber(cTimer.Elapsed / 1000, 3) & " seconds."
  
End Sub

Private Sub Command3_Click()

  cTimer.EndTimer 'Stop timer
  
  'Display elapsed time
  MsgBox "The timer was running for " & FormatNumber(cTimer.Elapsed, 0) & _
  " milliseconds, or " & FormatNumber(cTimer.Elapsed / 1000, 3) & " seconds."
  
  Command1.Enabled = True
  Command2.Enabled = False
  Command3.Enabled = False
  
End Sub
