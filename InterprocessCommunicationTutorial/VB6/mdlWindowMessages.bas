Attribute VB_Name = "mdlWindowMessaging"
Option Explicit

Private Const ModuleName = "mdlWindowsMessaging"


'***********************************
'      D E C L A R A T I O N S
'***********************************

' Declaration to register custom messages
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

' Declaration to create new window
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
  (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
  ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, _
  ByVal hInstance As Long, lpParam As Any) As Long

' Declaration to find window based on name and class type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_CLOSE = &H10

' Declaration to map MessageHandler to new function
Private Declare Function SetWindowLongApi Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Declaration to let MessageHandling fall through to original MessageHandler if it is not one of our custom messages
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Declaration to post async. message to target Window
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Integer) As Long

' Messages
Public Const MSG_HELLO_REQUEST = "Hello C#"
Public Const MSG_HELLO_RESPONSE = "Hello VB6"

' Window Titles
Private Const VB_WINDOWTITLE_SERVER = "prjInterprocessCommunication"
Private Const CS_WINDOWTITLE_SERVER = "InterprocessCommunication"

Private Potentials_WindowHandle As Long
Private WindowMessagingInitialised As Boolean

' New window handle
Private hWindowHandle As Long
' Old MessageHandler address: needed to reset MessageHandler back to original Handler
' else new window still listens to Messages inside VB6 IDE and causes crashes.
Private hOldProc As Long


'***********************************
'   P U B L I C    M E T H O D S
'***********************************

' Function to create custom window and setup Message Listening
Public Function InitWindowMessaging()

'This statement creates the new window
hWindowHandle = CreateWindowEx(0, "STATIC", VB_WINDOWTITLE_SERVER, 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)

' This statement sets the message handling to the ProcessWindowMessages function defined later.
' We also save the address (hOldProc) of the previous MessageHandler so we can reset on StopWindowMessaging
hOldProc = SetWindowLongApi(hWindowHandle, GWL_WNDPROC, AddressOf ProcessWindowMessages)

WindowMessagingInitialised = True

End Function


' Function to tear down Message Handling and return to original Message Handler
Public Function StopWindowMessaging()

' This statement sets the Message Handling to be set to the address of the previous MessageHandler which we saved before changing it to ours.
Call SetWindowLongApi(hWindowHandle, GWL_WNDPROC, hOldProc)

End Function


' Function to find C# window and attempt to send message
Public Function SendWindowsMessage(Message As String)
Dim hwndTarget As Long
Dim MessageId  As Long

If WindowMessagingInitialised = False Then InitWindowMessaging

' Get TargetWindow handle from global Window Name
hwndTarget = CSharp_WindowHandle

' Check if target window handler was found
If hwndTarget = 0 Then
    MsgBox "Unable to find the " & CS_WINDOWTITLE_SERVER & " window", vbCritical, "Interprocess Communication"
End If

' Get MessageId from API call to RegisterMessage
MessageId = HELLO_RESPONSE_MessageId

' If Window target exists, then SendMessage to target
If hwndTarget <> 0 Then Call PostMessage(hwndTarget, MessageId, 0, 0)
  
End Function


'***********************************
'  P R I V A T E   M E T H O D S
'***********************************

' Function to process messages. If not one of our custom messages, then fall through to original Message Handler
Private Function ProcessWindowMessages(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Integer) As Long

Select Case wMsg
    Case HELLO_REQUEST_MessageId
        MsgBox MSG_HELLO_REQUEST
        
    Case Else
        'Pass the message to the previous window procedure to handle it
        ProcessWindowMessages = CallWindowProc(hOldProc, hWnd, wMsg, wParam, lParam)
End Select

End Function


'***********************************
' P U B L I C   P R O P E R T I E S
'***********************************

' Uses API call to find window handle of C# application
Public Property Get CSharp_WindowHandle() As Long
CSharp_WindowHandle = FindWindow(vbNullString, CS_WINDOWTITLE_SERVER)
End Property

'Uses API call to create system-wide Message if not already created.
Public Property Get HELLO_RESPONSE_MessageId() As Long
Static HelloResponseMessageId As Long

'Pass in global Message_Name
If HelloResponseMessageId = 0 Then
    HelloResponseMessageId = RegisterWindowMessage(MSG_HELLO_RESPONSE)
End If

HELLO_RESPONSE_MessageId = HelloResponseMessageId
End Property

'Uses API call to create system-wide Message if not already created.
Public Property Get HELLO_REQUEST_MessageId() As Long
Static HelloRequestMessageId As Long

'Pass in global Message_Name
If HelloRequestMessageId = 0 Then
    HelloRequestMessageId = RegisterWindowMessage(MSG_HELLO_REQUEST)
End If

HELLO_REQUEST_MessageId = HelloRequestMessageId
End Property




