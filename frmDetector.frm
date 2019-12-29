VERSION 5.00
Begin VB.Form frmDetector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Device arrival / removal log"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11310
   Icon            =   "frmDetector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moDetector As DeviceDetector
Attribute moDetector.VB_VarHelpID = -1

Private Sub Form_Load()
  On Error Resume Next
  Set moDetector = New DeviceDetector
  
  If Not moDetector.StartMonitoring() Then
    MsgBox "Failed to register for notifications:" & vbCrLf & vbCrLf & moDetector.StartMonitoringErrorMsg, vbCritical
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  moDetector.EndMonitoring
  Set moDetector = Nothing
End Sub

Private Sub moDetector_OnDeviceArrival( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)
  
  Dim i             As Integer
  Dim sDriveLetter  As String
  Dim oDeviceInfo   As New DeviceInfo
  
  For i = 1 To piDriveCt
    sDriveLetter = Mid$(psDriveLetters, i, 1)
    List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                  InOutDeviceTypeToString(peDeviceType) & " " & _
                  "ARRIVAL #" & i & vbTab & _
                  sDriveLetter & vbTab & _
                  DriveTypeToString(paeDriveTypes(i))
    'Except for network drives, we can get more information for the device
    If peDeviceType <> eNetworkDevice Then
      If oDeviceInfo.GetDeviceInformation(sDriveLetter) Then
        List1.AddItem "  " & oDeviceInfo.VendorID & " / " & oDeviceInfo.ProductID & " #" & oDeviceInfo.SerialNumber
      Else
        List1.AddItem " Get device info error: " & oDeviceInfo.LastErrDesc
      End If
    End If
  Next i
End Sub

Private Sub moDetector_OnDeviceRemoval( _
  ByVal peDeviceType As eInOutDeviceType, _
  ByVal piDriveCt As Integer, _
  ByRef psDriveLetters As String, _
  ByRef paeDriveTypes() As eDriveType)
  
  Dim i   As Integer
  For i = 1 To piDriveCt
    List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                  InOutDeviceTypeToString(peDeviceType) & " " & _
                  "REMOVAL #" & i & vbTab & _
                  Mid$(psDriveLetters, i, 1) & vbTab & _
                  DriveTypeToString(paeDriveTypes(i))
  Next i
End Sub

Private Function InOutDeviceTypeToString(ByVal peDeviceType As eInOutDeviceType) As String
  Select Case peDeviceType
  Case eInOutDeviceType.ePhysicalDeviceOrDrive
    InOutDeviceTypeToString = "PHYSICAL"
  Case eInOutDeviceType.eMediaInDevice
    InOutDeviceTypeToString = "MEDIA   "
  Case eInOutDeviceType.eNetworkDevice
    InOutDeviceTypeToString = "NETWORK "
  Case Else
    InOutDeviceTypeToString = "UNKNOWN "
  End Select
End Function

Private Function DriveTypeToString(ByVal peDriveType As eDriveType) As String
  Select Case peDriveType
  Case eDriveType.eNoRootDir
    DriveTypeToString = "NOROOTDIR" '=removed or drive has no filesystem
  Case eDriveType.eRemoveable
    DriveTypeToString = "REMOVABLE"
  Case eDriveType.eFixed
    DriveTypeToString = "FIXED"
  Case eDriveType.eRemote
    DriveTypeToString = "REMOTE"
  Case eDriveType.eCDRom
    DriveTypeToString = "CDROM"
  Case eDriveType.eRamDisk
    DriveTypeToString = "RAMDISK"
  Case Else
    DriveTypeToString = "UNKNOWN"
  End Select
End Function


