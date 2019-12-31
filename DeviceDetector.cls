VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeviceDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'DeviceDetector class
'
'Responsibilities
'The DeviceDetector class watches for Windows device arrival and removal
'and notifies its owner of such events.
'
'How
'The work is delegated to a frmMain form instance that each instance of this class
'creates internally.
'frmMain notifies back the class instance when such an event occurs.
'The class instance then notifies its owner back either
'raising class events or invoking OLE callbacks.
'
' Notification method:
' -------------------------------+---------------------------------------------------------
' 1) Callback object is nothing  | Raise events OnDeviceArrival/OnDeviceRemoval
'    UseNoParamEvents is False   | (this is the default behavior)
' -------------------------------+---------------------------------------------------------
' 2) Callback object is nothing  | Raise events OnDeviceArrivalSignal/OnDeviceRemovalSignal
'    UseNoParamEvents is True    |
' -------------------------------+---------------------------------------------------------
' 3) Callback object is set      | Callback OnDeviceArrival/OnDeviceRemoval
'    UseNoParamEvents is False   |
' -------------------------------+---------------------------------------------------------
' 4) Callback object is set      | Callback OnDeviceArrivalSignal/OnDeviceRemovalSignal
'    UseNoParamEvents is True    |
' -------------------------------+---------------------------------------------------------
' For 2) and 4), class user has to call GetLastArrivalEventParams/GetLastRemovalEventParams
' to get the event parameters.
'
'Notes
'-----
'  These notes are valid for the ActiveX server project (AxDeviceDetector.vbp)
'  not for the standalone SaDeviceDetector.vbp, which is just a test driver.
'We instantiate a DeviceInfo object here, that we expose thru the DeviceInfoObject
'public property. There are three reasons to do that:
' 1. We know that there is a strong probability that class user will
'    need more information about the device being plugged/unplugged.
'    We simplify the task by providing an ready instance of DeviceInfo for that.
' 2. When we create the DeviceInfo class instance inside this ActiveX server,
'    it will run in the same thread as this class instance.
'    No need for the (external) class user to either instantiate a DeviceInfo object
'    outside the server (thus creating a new thread), or include the DeviceInfo
'    class in the outside project/host.
' 3. We have a dependency only on the 32bits deviceinfo.dll DLL, no need for
'    the 64bits DLL, even in 64bits hosts.
'    DLL project hosted here: https://github.com/francescofoti/DeviceInfo_dll

Private moDeviceInfo    As DeviceInfo

'If we set a callback object, the class will use late bound OLE callbacks
'on named methods (OnDeviceArrival, OnDeviceRemoval), instead of raising
'events.
Private moCallbackObject As Object

'we can change the behaviour of this class to raise events
'with no parameters changing this flag (defaults to False).
'We'll keep the parameters in member variables that the
'event sink object can call back while handling the events.
Private mfUseNoParamEvents As Boolean

'If we don't set a callback object, the class will try to send events
Event OnDeviceArrival( _
  ByVal psDeviceType As String, _
  ByVal piDriveCt As Integer, _
  ByVal psDriveLetters As String, _
  ByVal psDriveTypes As String)

Event OnDeviceRemoval( _
  ByVal psDeviceType As String, _
  ByVal piDriveCt As Integer, _
  ByVal psDriveLetters As String, _
  ByVal psDriveTypes As String)

'Events fired with no parameters (if mfUseNoParamEvents or not #ALLOW_PARAM_EVENTS)
Event OnDeviceArrivalSignal()
Event OnDeviceRemovalSignal()

'If we fire up events or OLE callbacks without parameters,
'we can then query for the event parameters (GetLast<xxx>EventParams methods),
'because we store them internally in these two member variables.
Private mtLastArrivalParams As TEventMessage
Private mtLastRemovalParams As TEventMessage

Private mfrmMain      As frmMain

'This class keeps track of the last error (after a public method call)
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

' Class events

Private Sub Class_Initialize()
  On Error Resume Next
  mfUseNoParamEvents = False
  Set moDeviceInfo = New DeviceInfo
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set moCallbackObject = Nothing
  Set moDeviceInfo = Nothing
End Sub

' Properties and methods

Public Property Get DeviceInfoObject() As DeviceInfo
  Set DeviceInfoObject = moDeviceInfo
End Property

Public Property Get UseNoParamEvents() As Boolean
  UseNoParamEvents = mfUseNoParamEvents
End Property

Public Property Let UseNoParamEvents(ByVal pfUseNoParamEvents As Boolean)
  mfUseNoParamEvents = pfUseNoParamEvents
End Property
  
Public Property Set CallbackObject(ByRef poCallbackObject As Object)
  Set moCallbackObject = poCallbackObject
End Property
  
Public Property Get CallbackObject() As Object
  Set CallbackObject = moCallbackObject
End Property

Public Sub GetLastArrivalEventParams( _
  ByRef psDeviceType As String, _
  ByRef piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef psDriveTypes As String)
  
  With mtLastArrivalParams
    psDeviceType = .sDeviceType
    piDriveCt = .iDriveCt
    psDriveLetters = .sDriveLetters
    psDriveTypes = .sDriveTypes
  End With
End Sub

Public Sub GetLastRemovalEventParams( _
  ByRef psDeviceType As String, _
  ByRef piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef psDriveTypes As String)

  With mtLastRemovalParams
    psDeviceType = .sDeviceType
    piDriveCt = .iDriveCt
    psDriveLetters = .sDriveLetters
    psDriveTypes = .sDriveTypes
  End With
End Sub

Friend Sub NotifyDeviceArrival( _
  ByVal psDeviceType As String, _
  ByVal piDriveCt As Integer, _
  ByVal psDriveLetters As String, _
  ByVal psDriveTypes As String)
  
  On Error Resume Next
  'DebugOutput "Device ARRIVAL: ct=" & piDriveCt & ", letters=" & psDriveLetters
  
  If mfUseNoParamEvents Then
    With mtLastArrivalParams
      .sDeviceType = psDeviceType
      .iDriveCt = piDriveCt
      .sDriveLetters = psDriveLetters
      .sDriveTypes = psDriveTypes
    End With
  End If
  
  If Not moCallbackObject Is Nothing Then
    'DebugOutput "Calling back method OnDeviceArrival(...)"
    If mfUseNoParamEvents Then
      'DebugOutput "Calling back method OnDeviceArrival(<no params>)"
      Call moCallbackObject.OnDeviceArrival
    Else
      'DebugOutput "Calling back method OnDeviceArrival(...)"
      Call moCallbackObject.OnDeviceArrival(psDeviceType, piDriveCt, psDriveLetters, psDriveTypes)
    End If
  Else
    If mfUseNoParamEvents Then
      'DebugOutput "Raising event OnDeviceArrivalSignal(<no params>)"
      RaiseEvent OnDeviceArrivalSignal
    Else
      'DebugOutput "Raising event OnDeviceArrival(...)"
      RaiseEvent OnDeviceArrival(psDeviceType, piDriveCt, psDriveLetters, psDriveTypes)
    End If
  End If
  
  'DebugOutput "NotifyDeviceArrival last error #&H" & Hex$(Err.Number) & ": " & Err.Description
End Sub

Friend Sub NotifyDeviceRemoval( _
  ByVal psDeviceType As String, _
  ByVal piDriveCt As Integer, _
  ByVal psDriveLetters As String, _
  ByVal psDriveTypes As String)
  
  On Error Resume Next
  'DebugOutput "Device REMOVAL: ct=" & piDriveCt & ", letters=" & psDriveLetters
  
  If mfUseNoParamEvents Then
    With mtLastRemovalParams
      .sDeviceType = psDeviceType
      .iDriveCt = piDriveCt
      .sDriveLetters = psDriveLetters
      .sDriveTypes = psDriveTypes
    End With
  End If
  
  If Not moCallbackObject Is Nothing Then
    If mfUseNoParamEvents Then
      'DebugOutput "Calling back method OnDeviceRemoval(<no params>)"
      Call moCallbackObject.OnDeviceRemoval
    Else
      'DebugOutput "Calling back method OnDeviceRemoval(...)"
      Call moCallbackObject.OnDeviceRemoval(psDeviceType, piDriveCt, psDriveLetters, psDriveTypes)
    End If
  Else
    If mfUseNoParamEvents Then
      'DebugOutput "Raising event OnDeviceRemovalSignal()"
      RaiseEvent OnDeviceRemovalSignal
    Else
      'DebugOutput "Raising event OnDeviceRemoval(...)"
      RaiseEvent OnDeviceRemoval(psDeviceType, piDriveCt, psDriveLetters, psDriveTypes)
    End If
  End If
  
  If Err.Number <> 0 Then
    'DebugOutput "NotifyDeviceRemoval Failed: " & Err.Description
  End If
End Sub

Public Function StartMonitoring() As Boolean
  Const LOCAL_ERR_CTX As String = "StartMonitoring"
  On Error GoTo StartMonitoring_Err
  Dim fOK As Boolean
  ClearErr
  
  Set mfrmMain = New frmMain
  Set mfrmMain.DetectorObject = Me
  fOK = mfrmMain.StartSubclassing()
  If Not fOK Then
    SetErr mfrmMain.LastErrCtx, mfrmMain.LastErr, mfrmMain.LastErrDesc
  End If
  StartMonitoring = fOK
  
StartMonitoring_Exit:
  Exit Function
StartMonitoring_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume StartMonitoring_Exit
End Function

Public Sub EndMonitoring()
  On Error Resume Next
  mfrmMain.StopSubclassing
  Set mfrmMain.DetectorObject = Nothing
  Unload mfrmMain
  Set mfrmMain = Nothing
End Sub

