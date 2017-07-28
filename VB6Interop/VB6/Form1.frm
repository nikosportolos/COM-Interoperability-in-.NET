VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   390
      TabIndex        =   0
      Top             =   630
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare a VB6Interop object
Public WithEvents VB6 As VB6Interop.VB6Interop
Attribute VB6.VB_VarHelpID = -1

Private Sub Command1_Click()

' Use SampleMethod() of VB6Interop
VB6.SampleMethod "Hello VB6!"

End Sub

Private Sub Form_Load()

' Initialize VB6Interop object
Set VB6 = New VB6Interop.VB6Interop

End Sub

' Implement SampleEvent
Private Sub VB6_SampleEvent(ByVal Message As String)
MsgBox Message
End Sub

