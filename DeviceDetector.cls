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

Public Enum eDriveType
  eNoRootDir = 1
  eRemoveable = 2
  eFixed = 3
  eRemote = 4
  eCDRom = 5
  eRamDisk = 6
End Enum

'Interpretation of dbcv_flags field of Win32 DEV_BROADCAST_VOLUME struct
'https://docs.microsoft.com/fr-fr/windows/win32/api/dbt/ns-dbt-dev_broadcast_volume
Public Enum eInOutDeviceType
  ePhysicalDeviceOrDrive
  eMediaInDevice
  eNetworkDevice
End Enum

Event OnDeviceArrival( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)

Event OnDeviceRemoval( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)

Private mfrmMain      As frmMain

'Notes:
' The goal for this project is to be an ActiveX EXE server, 1 thread per object.
' There's only one exposed object class, "DeviceDetector", ie this class.
' So, except if you're in the standalone version of this project:
' 1) there's error trapping everywhere here, we don't want the component to fail
'    on an exception (it would kill the ActiveX server process)
' 2) This is a component, throw NO UI around (frmMain stays hidden), even not a MsgBox.

Friend Sub NotifyDeviceArrival( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)
  
  On Error Resume Next
  RaiseEvent OnDeviceArrival(peDeviceType, piDriveCt, psDriveLetters, paeDriveTypes())
End Sub

Friend Sub NotifyDeviceRemoval( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)
  
  On Error Resume Next
  RaiseEvent OnDeviceRemoval(peDeviceType, piDriveCt, psDriveLetters, paeDriveTypes())
End Sub

Public Function StartMonitoring() As Boolean
  On Error Resume Next
  Set mfrmMain = New frmMain
  Set mfrmMain.DetectorObject = Me
  StartMonitoring = mfrmMain.StartSubclassing()
End Function

'If something goes wrong with the Win32 API, ie when starting the subclaasing or
'the notification registration tasks, we'll get the error message from the Win32 API.
Public Function StartMonitoringErrorMsg() As String
  On Error Resume Next
  StartMonitoringErrorMsg = LastDllErrorMsg()
End Function

Public Sub EndMonitoring()
  On Error Resume Next
  mfrmMain.StopSubclassing
  Set mfrmMain.DetectorObject = Nothing
  Unload mfrmMain
  Set mfrmMain = Nothing
End Sub
