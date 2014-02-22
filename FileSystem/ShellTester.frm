VERSION 5.00
Begin VB.Form frmShellTester 
   Caption         =   "Shell Tester"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTimeout 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Text            =   "-1"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtCommand 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Notepad"
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdResponsize 
      Caption         =   "Remain Responsive"
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdFreeze 
      Caption         =   "Freeze Calling Window"
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "command"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "timeout (s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   825
   End
End
Attribute VB_Name = "frmShellTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdResponsize_Click()
    Call F_.ShellWait(txtCommand.Text _
                    , False _
                    , _
                    , vbNormalFocus)
    
    MsgBox "Execution resumed."
    
End Sub

Private Sub cmdFreeze_Click()
    Call F_.ShellWait(txtCommand.Text _
                    , True _
                    , CLng(txtTimeout.Text) * 1000 _
                    , vbNormalFocus)
    
    MsgBox "Execution resumed."
    
End Sub

