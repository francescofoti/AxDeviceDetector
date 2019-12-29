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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Implements ISubclass

Private mfNeedsToUnregisterNotifications As Boolean
Private mfIsSubclassing As Boolean

Private Sub Form_Unload(Cancel As Integer)
  StopSubclassing
End Sub

Public Function StartSubclassing() As Boolean
  If Not mfIsSubclassing Then
    SubclassWindowForNotifications Me, Me.hWnd
    mfIsSubclassing = True
  End If
  mfNeedsToUnregisterNotifications = RegisterForDevicesNotifications(Me.hWnd, mhDevNotify)
  StartSubclassing = mfNeedsToUnregisterNotifications
End Function

Public Sub StopSubclassing()
  If mfNeedsToUnregisterNotifications Then Call UnRegisterForDevicesNotifications(mhDevNotify)
  If mfIsSubclassing Then Call UnsubclassWindowForNotifications(Me.hWnd)
End Sub

'Will receive only WM_DEVICECHANGE messages
'MSDN: https://docs.microsoft.com/en-us/windows/win32/devio/detecting-media-insertion-or-removal
Private Function ISubclass_WindowProc(ByVal plhWnd As Long, ByVal plMsg As Long, ByVal pwParam As Long, ByVal plParam As Long) As Long
  Dim tDevBroadcastHeader As DEV_BROADCAST_HDR
  Dim lUnitMask           As Long
  Dim iFlags              As Integer
  Dim DeviceGUID          As Guid
  Dim lpDeviceName        As Long
  Dim iDrives             As Integer
  Dim iDriveCt            As Integer
  'Plausible reentrancy, so keep the following variables as stack variables, don't scope them outside this function
  Dim eDeviceType         As eInOutDeviceType
  Dim sDriveLetters       As String
  Dim aeDriveTypes(1 To MAX_DRIVES) As eDriveType
  
  If plhWnd <> Me.hWnd Then Exit Function 'Could be just asserted, anyway something would be really wrong
  
  On Error Resume Next 'no runtime exceptions from here or we'll crash
  
  Select Case pwParam
  
  Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
    If (plParam) Then ' Read generic DEV_BROADCAST_HDR structure
      Call RtlMoveMemory(tDevBroadcastHeader, ByVal plParam, Len(tDevBroadcastHeader))
      
      If (tDevBroadcastHeader.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
        ' Specific information is after the header structure
        Call RtlMoveMemory(lUnitMask, ByVal (plParam + Len(tDevBroadcastHeader)), 4&)
        Call RtlMoveMemory(iFlags, ByVal (plParam + Len(tDevBroadcastHeader) + 4&), 2&)
        
        iDriveCt = DrivesFromMask(lUnitMask, sDriveLetters, aeDriveTypes())
        eDeviceType = iFlags
        
        'Notify the event
        If Not Me.DetectorObject Is Nothing Then
          If pwParam = DBT_DEVICEARRIVAL Then
            Call DetectorObject.NotifyDeviceArrival(eDeviceType, iDriveCt, sDriveLetters, aeDriveTypes())
          Else
            Call DetectorObject.NotifyDeviceRemoval(eDeviceType, iDriveCt, sDriveLetters, aeDriveTypes())
          End If
        End If
      End If
    End If
  
  End Select
End Function

Private Function DrivesFromMask(ByVal plUnitMask As Long, ByRef psDriveLetters As String, ByRef paeDriveTypes() As eDriveType) As Integer
  Dim i             As Integer
  Dim iDriveCt      As Integer
  Dim sDrive        As String
  
  Debug.Assert (LBound(paeDriveTypes) = 1) And (UBound(paeDriveTypes) = MAX_DRIVES)
  
  psDriveLetters = Space$(MAX_DRIVES)
  For i = 0 To MAX_DRIVES - 1
    If plUnitMask And (2 ^ i) Then
      iDriveCt = iDriveCt + 1
      sDrive = Chr$(Asc("A") + i)
      Mid$(psDriveLetters, iDriveCt, 1) = sDrive
      sDrive = sDrive & ":\"
      paeDriveTypes(iDriveCt) = GetDriveType(sDrive)
    End If
  Next i
  DrivesFromMask = iDriveCt
End Function

