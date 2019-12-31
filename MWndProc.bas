Attribute VB_Name = "MWndProc"
Option Explicit

Public Declare Function RegisterDeviceNotification Lib "User32.dll" Alias _
                         "RegisterDeviceNotificationA" (ByVal phRecipient As Long, _
                          ByRef NotificationFilter As Any, ByVal plflags As Long) As Long
Public Declare Function UnregisterDeviceNotification Lib "User32.dll" ( _
                          ByVal plhWnd As Long) As Long

Public Const DEVICE_NOTIFY_WINDOW_HANDLE  As Long = &H0&
Public Const WM_DEVICECHANGE              As Long = &H219&

Public Const DBT_DEVNODES_CHANGED       As Long = &H7&
Public Const DBT_DEVICEARRIVAL          As Long = &H8000&
Public Const DBT_DEVICEREMOVECOMPLETE   As Long = &H8004&
Public Const DBT_DEVTYP_VOLUME          As Long = &H2&      ' Logical volume
Public Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5&      ' Device interface class

Public Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Type DEV_BROADCAST_DEVICEINTERFACE
  dbcc_size As Long
  dbcc_devicetype As Long
  dbcc_reserved As Long
  dbcc_classguid As Guid
  dbcc_name As Long
End Type

Public Type DEV_BROADCAST_HDR
  dbch_size As Long
  dbch_devicetype As Long
  dbch_reserved As Long
End Type

'Windows API used only in this module
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000&
Private Declare Function GetLastError& Lib "kernel32" ()
Private Declare Function FormatMessageW Lib "kernel32" (ByVal pdwFlags As Long, ByVal plSource As Long, ByVal pdwMessageId As Long, ByVal pdwLanguageId As Long, ByVal plBuffer As Long, ByVal plSize As Long, plArguments As Long) As Long

'Variables needed for subclassing
'WARNING: this code cannot handle multiple windows subclassing:
' - there's only one mlOldWindowProcAddress and miiSubclass module level variables
' - To fix that, use something like a collection, indexed on the hWnd
Private miiSubclass As ISubclass
Private mlOldWindowProcAddress    As Long

Private Const GWL_WNDPROC As Long = -4&
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal plOldWndProc As Long, ByVal plhWnd As Long, ByVal plMsg As Long, ByVal pwParam As Long, ByVal plParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal plhWnd As Long, ByVal plPropIndex As Long, ByVal pdwNewValue As Long) As Long

'We'll repack the event messages in this
'structure, so we can queue them in a CEventQueue class,
'to notify them later via ActiveX events or OLE callbacks;
'we can't notify out of process while processing the window message.
Public Const EVENTID_ARRIVAL As Integer = 0
Public Const EVENTID_REMOVAL As Integer = 1
Public Type TEventMessage
  iEventID      As Integer
  sDeviceType   As String
  iDriveCt      As Integer
  sDriveLetters As String
  sDriveTypes   As String
End Type

#If Win64 Then
private Declare PtrSafe Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
#Else
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
#End If

'Send the message to the Windows debug monitor.
'To see the debug messages, uncomment them in the source code,
'download Mark Russinovitch's DebugView from sysinternals
' https://docs.microsoft.com/fr-fr/sysinternals/
'run it with administrator privileges and check "Capture Win32"
'and also "Capture global Win32" in the capture menu.
Public Sub DebugOutput(ByVal psMessage As String)
  On Error Resume Next
  OutputDebugString psMessage
End Sub

#If STANDALONE_VERSION Then
  Public Sub Main()
    frmDetector.Show
  End Sub
#Else
  Public Sub Main()
  End Sub
#End If

Public Sub SubclassWindowForNotifications(ByRef piiSubclass As ISubclass, ByVal plhWnd As Long)
  'Subclass this window so we can handle notification messages
  Set miiSubclass = piiSubclass
  mlOldWindowProcAddress = SetWindowLong(plhWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Function WindowProc(ByVal plhWnd As Long, ByVal plMsg As Long, ByVal pwParam As Long, ByVal plParam As Long) As Long
  Dim lResult As Long
  
  If miiSubclass Is Nothing Then Exit Function 'foolprof just to avoid a crash
  
  'We're only interested in the WM_WTSSESSION_CHANGE message
  If plMsg = WM_DEVICECHANGE Then
    On Error Resume Next
    lResult = miiSubclass.WindowProc(plhWnd, plMsg, pwParam, plParam)
  End If
  'Call previous wndproc for every message, let it have the last word on the return
  'value, as we do not respond to message that need to return smthg <> 0
  lResult = CallWindowProc(mlOldWindowProcAddress, plhWnd, plMsg, pwParam, plParam)
  
  WindowProc = lResult
End Function

Public Sub UnsubclassWindowForNotifications(ByVal plhWnd As Long)
  'Unsubclass window
  Call SetWindowLong(plhWnd, GWL_WNDPROC, mlOldWindowProcAddress)
  Set miiSubclass = Nothing
End Sub

Public Function RegisterForDevicesNotifications(ByVal plhWnd As Long, ByRef plRethDevNotify As Long) As Boolean
  Dim tNotificationFilter As DEV_BROADCAST_DEVICEINTERFACE
  
  With tNotificationFilter
    .dbcc_size = Len(tNotificationFilter)
    .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
  End With
  
  plRethDevNotify = RegisterDeviceNotification(plhWnd, tNotificationFilter, DEVICE_NOTIFY_WINDOW_HANDLE)
  
  'If this fails, when can use LastDllErrorMsg() to get error information
  RegisterForDevicesNotifications = CBool(plRethDevNotify <> 0)
End Function

Public Sub UnRegisterForDevicesNotifications(ByVal plhDevNotify As Long)
  Call UnregisterDeviceNotification(plhDevNotify)
End Sub

Public Function LastDllErrorMsg(Optional ByVal plErrCode As Long) As String
  Dim sBuffer         As String   ' Place where error description will be copied to.
  Dim lCopiedCt       As Long     ' Number of bytes copied to sBuffer
  Const BUFFER_SIZE As Long = 2048&
  
  If plErrCode = 0 Then                       ' no error code supplied
    plErrCode = Err.LastDllError              ' use the VB last known API error code
  Else
    plErrCode = Abs(plErrCode)                ' user supplied DLL error code
  End If
  If plErrCode = 0 Then Exit Function         ' bail if no error code
  ' prepare the buffer
  sBuffer = Space$(BUFFER_SIZE)
  ' translate the error code
  lCopiedCt = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, _
                         0&, plErrCode, _
                         0&, StrPtr(sBuffer), BUFFER_SIZE - 1&, 0&)
  LastDllErrorMsg = Trim$(sBuffer) 'trailing \0 not removed here
End Function

