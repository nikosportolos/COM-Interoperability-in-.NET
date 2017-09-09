VERSION 5.00
Begin VB.Form frmInteropVB6 
   Caption         =   "COM Interoperability in .NET"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "frmInteropVB6"
   ScaleHeight     =   2160
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCurrentTime 
      Height          =   585
      Left            =   3390
      TabIndex        =   3
      Top             =   1200
      Width           =   2265
   End
   Begin VB.CommandButton cmdGetTime 
      Caption         =   "Get Current Time"
      Height          =   585
      Left            =   870
      TabIndex        =   2
      Top             =   1200
      Width           =   2265
   End
   Begin VB.TextBox txtSampleMethod 
      Height          =   585
      Left            =   3390
      TabIndex        =   1
      Top             =   510
      Width           =   2265
   End
   Begin VB.CommandButton cmdTestSampleMethod 
      Caption         =   "Test Sample Method"
      Height          =   585
      Left            =   870
      TabIndex        =   0
      Top             =   495
      Width           =   2265
   End
End
Attribute VB_Name = "frmInteropVB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare a VB6Interop object
Public WithEvents VB6 As VB6Interop.VB6Interop
Attribute VB6.VB_VarHelpID = -1

Private Sub cmdGetTime_Click()
Dim sTime As String

sTime = VB6.GetTime
Me.txtCurrentTime.Text = sTime

End Sub

Private Sub cmdTestSampleMethod_Click()

' Use SampleMethod() of VB6Interop
''VB6.SampleMethod "Hello VB6!"

If (Len(Me.txtSampleMethod.Text) = 0) Then
    MsgBox "Sample text cannot be empty!", vbCritical, "COM Interoperability in .NET"
Else
    Call VB6.SampleMethod(Me.txtSampleMethod.Text)
End If


End Sub

Private Sub Form_Load()

' Initialize VB6Interop object
Set VB6 = New VB6Interop.VB6Interop

End Sub

' Implement SampleEvent
Private Sub VB6_SampleEvent(ByVal Message As String)
MsgBox Message, vbOKOnly, "COM Interoperability in .NET"
End Sub

