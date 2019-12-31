VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Device arrival / removal detector"
   ClientHeight    =   270
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timNotify 
      Interval        =   200
      Left            =   480
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'https://jeffpar.github.io/kbarchive/kb/190/Q190523/
'
'NOTE: This window is never shown, it stays loaded, but invisible

'API types used here only
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
                      ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'This object will be notified of session lock/unlock detection, so that it can raise an event
Public DetectorObject As DeviceDetector

'When we register for device notifications messages,
'the Windows API gives us back as notification handle (a Long)
Private mhDevNotify   As Long

Private Const MAX_DRIVES  As Integer = 26 'A to Z

'We queue the messages we get thru ISubclass to notify them later in the Form's Timer
Private Const NOTIFY_INTERVAL As Integer = 200  'millisecons
Private Const MAX_QUEUE_CAPACITY As Integer = 50
Private moEventQueue  As CEventQueue

Implements ISubclass

Private mfNeedsToUnregisterNotifications As Boolean
Private mfIsSubclassing As Boolean

Private mlErrNo   As Long
Private msErrCtx  As String
Private msErrDesc As String

Private Sub ClearErr()
  mlErrNo = 0&
  msErrCtx = ""
  msErrDesc = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErrNum As Long, ByVal psErrDesc As String)
  mlErrNo = plErrNum
  msErrCtx = psErrCtx
  msErrDesc = psErrDesc
  'DebugOutput App.Title & " Error, context:" & msErrCtx & ", #" & mlErrNo & ": " & msErrDesc
End Sub

Public Property Get LastErr() As Long
  LastErr = mlErrNo
End Property

Public Property Get LastErrDesc() As String
  LastErrDesc = msErrDesc
End Property

Public Property Get LastErrCtx() As String
  LastErrCtx = msErrCtx
End Property

Private Sub Form_Load()
  Const LOCAL_ERR_CTX As String = "Form_Load"
  On Error GoTo Load_Err
  Dim fOK       As Boolean
  ClearErr
  
  Set moEventQueue = New CEventQueue
  fOK = moEventQueue.CreateQueue(MAX_QUEUE_CAPACITY)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, moEventQueue.LastErr, "Failed to init message queue for #" & MAX_QUEUE_CAPACITY & " messages: " & moEventQueue.LastErrDesc
  End If
  
Load_Exit:
  Exit Sub

Load_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Load_Exit
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  StopSubclassing
  timNotify.Enabled = False
End Sub

Public Function StartSubclassing() As Boolean
  If Not mfIsSubclassing Then
    SubclassWindowForNotifications Me, Me.hWnd
    mfIsSubclassing = True
  End If
  mfNeedsToUnregisterNotifications = RegisterForDevicesNotifications(Me.hWnd, mhDevNotify)
  If Not mfNeedsToUnregisterNotifications Then
    SetErr "StartSubclassing", -1&, LastDllErrorMsg()
  End If
  StartSubclassing = mfNeedsToUnregisterNotifications
End Function

Public Sub StopSubclassing()
  If mfNeedsToUnregisterNotifications Then Call UnRegisterForDevicesNotifications(mhDevNotify)
  If mfIsSubclassing Then Call UnsubclassWindowForNotifications(Me.hWnd)
End Sub

Private Function ActivatePushNotificationTimer() As Boolean
  On Error GoTo ActivatePushNotificationTimer_Err
  
  'DebugOutput "Activating notification timer"
  'Just activate the timer and let it run, if not already so
  If Not timNotify.Enabled Then
    timNotify.Interval = NOTIFY_INTERVAL
    timNotify.Enabled = True
  End If
  
  ActivatePushNotificationTimer = True
  
ActivatePushNotificationTimer_Exit:
  Exit Function

ActivatePushNotificationTimer_Err:
  'DebugOutput "Failed to activate the timer: " & Err.Description
  SetErr "ActivatePushNotificationTimer", Err.Number, Err.Description
  Resume ActivatePushNotificationTimer_Exit
End Function

'Will receive only WM_DEVICECHANGE messages
'MSDN: https://docs.microsoft.com/en-us/windows/win32/devio/detecting-media-insertion-or-removal
Private Function ISubclass_WindowProc(ByVal plhWnd As Long, ByVal plMsg As Long, ByVal pwParam As Long, ByVal plParam As Long) As Long
  Const LOCAL_ERR_CTX As String = "WindowProc"
  Dim tDevBroadcastHeader As DEV_BROADCAST_HDR
  Dim lUnitMask           As Long
  Dim iFlags              As Integer
  Dim DeviceGUID          As Guid
  Dim lpDeviceName        As Long
  Dim iDrives             As Integer
  Dim iEventID            As Integer
  Dim oMessage            As New CEventMessage
  Dim sDriveLetters       As String
  Dim sDriveTypes         As String
  Dim fOK                 As Boolean
  
  If plhWnd <> Me.hWnd Then Exit Function 'Could be just asserted, anyway something would be really wrong
  
  On Error Resume Next 'no runtime exceptions from here or we'll crash
  ClearErr
  
  Select Case pwParam
  
  Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
    'DebugOutput "AxDeviceDetector: DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE"
    If (plParam) Then ' Read generic DEV_BROADCAST_HDR structure
      Call RtlMoveMemory(tDevBroadcastHeader, ByVal plParam, Len(tDevBroadcastHeader))
      
      If (tDevBroadcastHeader.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
        ' Specific information is after the header structure
        Call RtlMoveMemory(lUnitMask, ByVal (plParam + Len(tDevBroadcastHeader)), 4&)
        Call RtlMoveMemory(iFlags, ByVal (plParam + Len(tDevBroadcastHeader) + 4&), 2&)
        
        oMessage.DriveCt = DrivesFromMask(lUnitMask, sDriveLetters, sDriveTypes)
        oMessage.DriveLetters = sDriveLetters
        oMessage.DriveTypes = sDriveTypes
        'Flags to a meaningful string
        Select Case iFlags
        Case 0
          oMessage.DeviceType = "PHYSICAL"
        Case 1
          oMessage.DeviceType = "MEDIA"
        Case 2
          oMessage.DeviceType = "NETWORK"
        End Select
        
        'Push the event in the event queue, the timer will handle it
        If Not Me.DetectorObject Is Nothing Then
          If pwParam = DBT_DEVICEARRIVAL Then
            oMessage.EventID = EVENTID_ARRIVAL
          Else
            oMessage.EventID = EVENTID_REMOVAL
          End If
          'DebugOutput "Pushing message, EventID #" & oMessage.EventID & _
                      ",letters: " & Trim$(oMessage.DriveLetters) & _
                      ",types:" & Trim$(oMessage.DriveTypes) & _
                      ",device: " & oMessage.DeviceType
          fOK = moEventQueue.QPush(oMessage)
          If fOK Then
            'start the timer
            Call ActivatePushNotificationTimer
          Else
            SetErr LOCAL_ERR_CTX, moEventQueue.LastErr, moEventQueue.LastErrDesc
            'message lost, queue full, clear the whole queue to keep latest messages
            moEventQueue.Clear
          End If
        End If
      End If
    End If
  
  End Select
End Function

Private Function DrivesFromMask(ByVal plUnitMask As Long, ByRef psDriveLetters As String, ByRef psDriveTypes As String) As Integer
  Dim i             As Integer
  Dim iDriveCt      As Integer
  Dim sDrive        As String
  
  psDriveLetters = Space$(MAX_DRIVES)
  psDriveTypes = Space$(MAX_DRIVES)
  For i = 0 To MAX_DRIVES - 1
    If plUnitMask And (2 ^ i) Then
      iDriveCt = iDriveCt + 1
      sDrive = Chr$(Asc("A") + i)
      Mid$(psDriveLetters, iDriveCt, 1) = sDrive
      sDrive = sDrive & ":\"
      Mid$(psDriveTypes, iDriveCt, 1) = Chr$(Asc("0") + GetDriveType(sDrive))
    End If
  Next i
  DrivesFromMask = iDriveCt
End Function

Private Sub timNotify_Timer()
  On Error Resume Next
  'DebugOutput "Timer event"
  
  'Pop and notify detector object
  If Me.DetectorObject Is Nothing Then
    'DebugOutput "Timer: No detector object"
    Exit Sub
  End If
  If moEventQueue Is Nothing Then
    'DebugOutput "Timer: No event queue"
    Exit Sub
  End If
  If moEventQueue.Count = 0 Then
    'DebugOutput "Timer: Queue is empty"
    timNotify.Enabled = False
    Exit Sub
  End If
  
  Dim oMessage As New CEventMessage
  Dim fOK      As Boolean
  
  ClearErr
  fOK = moEventQueue.QPop(oMessage)
  If fOK Then
    'DebugOutput "Notifying detector object"
    If oMessage.EventID = EVENTID_ARRIVAL Then
      Call DetectorObject.NotifyDeviceArrival(oMessage.DeviceType, oMessage.DriveCt, oMessage.DriveLetters, oMessage.DriveTypes)
    Else
      Call DetectorObject.NotifyDeviceRemoval(oMessage.DeviceType, oMessage.DriveCt, oMessage.DriveLetters, oMessage.DriveTypes)
    End If
  Else
    SetErr "NotifyTimer", moEventQueue.LastErr, moEventQueue.LastErrDesc
  End If
End Sub
