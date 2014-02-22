VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SelfTimer demo"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Timer As SelfTimer
Attribute Timer.VB_VarHelpID = -1

Private Sub Form_Load()
    Set Timer = New SelfTimer
    Timer.Interval = 1
End Sub

Private Sub Form_Terminate()
    Set Timer = Nothing
End Sub

Private Sub Timer_Timer(ByVal Seconds As Currency)
    lblTimer.Caption = Format$(Seconds, "0.000") & " seconds has passed"
End Sub
