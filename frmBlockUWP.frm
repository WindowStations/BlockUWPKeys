VERSION 5.00
Begin VB.Form frmBlockUWP 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   Icon            =   "frmBlockUWP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrHook 
      Interval        =   400
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmBlockUWP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ispainted As Boolean
Private Sub Form_Load()
   modKeyboard.HookKeyboard
End Sub
Private Sub Form_Paint()
   If ispainted = True Then Exit Sub
   ispainted = True
   Me.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
   modKeyboard.UnhookKeyboard
End Sub
Private Sub tmrHook_Timer()
   modKeyboard.HookKeyboard
End Sub
Private Sub Command1_Click()
   Unload Me
End Sub
