VERSION 5.00
Begin VB.Form frmInterComm 
   Caption         =   "VB6InterComm"
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetWindowTitle 
      Caption         =   "Get window title"
      Height          =   525
      Left            =   2805
      TabIndex        =   1
      Top             =   540
      Width           =   2025
   End
   Begin VB.CommandButton cmdTestMessage 
      Caption         =   "Send Test Message"
      Height          =   525
      Left            =   585
      TabIndex        =   0
      Top             =   540
      Width           =   2025
   End
End
Attribute VB_Name = "frmInterComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetWindowTitle_Click()
Call mdlWindowMessaging.GetWindowTitle(Me.hWnd)
End Sub

Private Sub cmdTestMessage_Click()
'Call mdlWindowMessaging.SendWindowsMessage
Call mdlWindowMessaging.SendMessageToCSharp
End Sub

Private Sub Form_Load()
'Create custom window and start listening to window messages.
Call mdlWindowMessaging.InitWindowMessaging
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Tear down custom message handling and pass back 
'to original message handler i.e this form
Call mdlWindowMessaging.StopWindowMessaging
End Sub
