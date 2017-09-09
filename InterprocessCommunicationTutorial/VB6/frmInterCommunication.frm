VERSION 5.00
Begin VB.Form frmInterComm 
   Caption         =   "Interprocess Communication Tutorial"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmInterComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Create custom window and start listening to window messages.
mdlWindowMessaging.InitWindowMessaging
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Tear down custom message handling and pass back to original message handler i.e this form
mdlWindowMessaging.StopWindowMessaging
End Sub
