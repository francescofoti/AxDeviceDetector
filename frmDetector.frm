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

#If Not STANDALONE_VERSION Then
Private WithEvents moDetector As AxDeviceDetector.DeviceDetector
Attribute moDetector.VB_VarHelpID = -1
#Else
Private WithEvents moDetector As DeviceDetector
Attribute moDetector.VB_VarHelpID = -1
#End If

Private Sub Form_Load()
  On Error Resume Next
  Set moDetector = New DeviceDetector
  'Doesn't work, we get automation errors:
  'Set moDetector.CallbackObject = Me
  'moDetector.UseNoParamEvents = True
  
  If Not moDetector.StartMonitoring() Then
    MsgBox "Failed to register for notifications:" & vbCrLf & vbCrLf & moDetector.LastErrDesc, vbCritical
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  moDetector.EndMonitoring
  Set moDetector = Nothing
End Sub

#If Not NOTIFY_WITH_CALLBACKS Then
  #If NOTIFY_WITH_PARAMS Then
    
    Private Sub moDetector_OnDeviceArrival(ByVal psDeviceType As String, ByVal piDriveCt As Integer, ByVal psDriveLetters As String, ByVal psDriveTypes As String)
      Dim i             As Integer
      Dim sDriveLetter  As String
      
      For i = 1 To piDriveCt
        sDriveLetter = Mid$(psDriveLetters, i, 1)
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & "!" & _
                      psDeviceType & " " & _
                      "ARRIVAL #" & i & vbTab & _
                      sDriveLetter & vbTab & _
                      DriveTypeToString(Mid$(psDriveTypes, i, 1))
        'Except for network drives, we can get more information for the device
        If psDeviceType <> "NETWORK" Then
          If moDetector.DeviceInfoObject.GetDeviceInformation(sDriveLetter) Then
            List1.AddItem "  " & moDetector.DeviceInfoObject.VendorID & " / " & moDetector.DeviceInfoObject.ProductID & " #" & moDetector.DeviceInfoObject.SerialNumber
          Else
            List1.AddItem " Get device info error: " & moDetector.DeviceInfoObject.LastErrDesc
          End If
        End If
      Next i
    End Sub
  
    Private Sub moDetector_OnDeviceRemoval(ByVal psDeviceType As String, ByVal piDriveCt As Integer, ByVal psDriveLetters As String, ByVal psDriveTypes As String)
      Dim i   As Integer
      For i = 1 To piDriveCt
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      psDeviceType & " " & _
                      "REMOVAL #" & i & vbTab & _
                      Mid$(psDriveLetters, i, 1) & vbTab & _
                      DriveTypeToString(Mid$(psDriveTypes, i, 1))
      Next i
    End Sub
    
  #Else
    Private Sub moDetector_OnDeviceArrivalSignal()
      Dim sDeviceType   As String
      Dim iDriveCt      As Integer
      Dim sDriveLetters As String
      Dim sDriveTypes   As String
      
      Dim i             As Integer
      Dim sDriveLetter  As String
      
      moDetector.GetLastArrivalEventParams sDeviceType, iDriveCt, sDriveLetters, sDriveTypes
      
      For i = 1 To iDriveCt
        sDriveLetter = Mid$(sDriveLetters, i, 1)
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      sDeviceType & " " & _
                      "ARRIVAL #" & i & vbTab & _
                      sDriveLetter & vbTab & _
                      DriveTypeToString(Mid$(sDriveTypes, i, 1))
        'Except for network drives, we can get more information for the device
        If sDeviceType <> "NETWORK" Then
          If moDetector.DeviceInfoObject.GetDeviceInformation(sDriveLetter) Then
            List1.AddItem "  " & moDetector.DeviceInfoObject.VendorID & " / " & moDetector.DeviceInfoObject.ProductID & " #" & moDetector.DeviceInfoObject.SerialNumber
          Else
            List1.AddItem " Get device info error: " & moDetector.DeviceInfoObject.LastErrDesc
          End If
        End If
      Next i
    End Sub

    Private Sub moDetector_OnDeviceRemovalSignal()
      Dim sDeviceType   As String
      Dim iDriveCt      As Integer
      Dim sDriveLetters As String
      Dim sDriveTypes   As String
      
      moDetector.GetLastRemovalEventParams sDeviceType, iDriveCt, sDriveLetters, sDriveTypes
    
      Dim i   As Integer
      For i = 1 To iDriveCt
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      sDeviceType & " " & _
                      "REMOVAL #" & i & vbTab & _
                      Mid$(sDriveLetters, i, 1) & vbTab & _
                      DriveTypeToString(Mid$(sDriveTypes, i, 1))
      Next i
    End Sub

  #End If
#End If

#If NOTIFY_WITH_CALLBACKS Then
  #If NOTIFY_WITH_PARAMS Then
    Public Sub OnDeviceArrival( _
      ByVal psDeviceType As String, _
      ByVal piDriveCt As Integer, _
      ByVal psDriveLetters As String, _
      ByVal psDriveTypes As String)
      
      Dim i             As Integer
      Dim sDriveLetter  As String
      
      For i = 1 To piDriveCt
        sDriveLetter = Mid$(psDriveLetters, i, 1)
        Debug.Print Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      psDeviceType & " " & _
                      "ARRIVAL #" & i & vbTab & _
                      sDriveLetter & vbTab & _
                      DriveTypeToString(Mid$(psDriveTypes, i, 1))
        'Except for network drives, we can get more information for the device
        If psDeviceType <> "NETWORK" Then
          If moDetector.DeviceInfoObject.GetDeviceInformation(sDriveLetter) Then
            Debug.Print "  " & moDetector.DeviceInfoObject.VendorID & " / " & moDetector.DeviceInfoObject.ProductID & " #" & moDetector.DeviceInfoObject.SerialNumber
          Else
            Debug.Print " Get device info error: " & moDetector.DeviceInfoObject.LastErrDesc
          End If
        End If
      Next i
    End Sub
    
    Public Sub OnDeviceRemoval( _
      ByVal psDeviceType As String, _
      ByVal piDriveCt As Integer, _
      ByVal psDriveLetters As String, _
      ByVal psDriveTypes As String)
      
      Dim i   As Integer
      For i = 1 To piDriveCt
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      psDeviceType & " " & _
                      "REMOVAL #" & i & vbTab & _
                      Mid$(psDriveLetters, i, 1) & vbTab & _
                      DriveTypeToString(Mid$(psDriveTypes, i, 1))
      Next i
    End Sub
  
  #Else
    Public Sub OnDeviceArrival()
      Dim sDeviceType   As String
      Dim iDriveCt      As Integer
      Dim sDriveLetters As String
      Dim sDriveTypes   As String
      
      Dim i             As Integer
      Dim sDriveLetter  As String
      
      moDetector.GetLastArrivalEventParams sDeviceType, iDriveCt, sDriveLetters, sDriveTypes
      
      For i = 1 To iDriveCt
        sDriveLetter = Mid$(sDriveLetters, i, 1)
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      sDeviceType & " " & _
                      "ARRIVAL #" & i & vbTab & _
                      sDriveLetter & vbTab & _
                      DriveTypeToString(Mid$(sDriveTypes, i, 1))
        'Except for network drives, we can get more information for the device
        If sDeviceType <> "NETWORK" Then
          If moDetector.DeviceInfoObject.GetDeviceInformation(sDriveLetter) Then
            List1.AddItem "  " & moDetector.DeviceInfoObject.VendorID & " / " & moDetector.DeviceInfoObject.ProductID & " #" & moDetector.DeviceInfoObject.SerialNumber
          Else
            List1.AddItem " Get device info error: " & moDetector.DeviceInfoObject.LastErrDesc
          End If
        End If
      Next i
    End Sub
    
    Public Sub OnDeviceRemoval()
      Dim sDeviceType   As String
      Dim iDriveCt      As Integer
      Dim sDriveLetters As String
      Dim sDriveTypes   As String
      
      moDetector.GetLastRemovalEventParams sDeviceType, iDriveCt, sDriveLetters, sDriveTypes
    
      Dim i   As Integer
      For i = 1 To iDriveCt
        List1.AddItem Format$(Now, "dd.mm.yyyy hh:mm:ss") & " " & _
                      sDeviceType & " " & _
                      "REMOVAL #" & i & vbTab & _
                      Mid$(sDriveLetters, i, 1) & vbTab & _
                      DriveTypeToString(Mid$(sDriveTypes, i, 1))
      Next i
    End Sub
  
  #End If
#End If

Private Function DriveTypeToString(ByVal psDriveType As String) As String
  Select Case psDriveType
  Case "1"
    DriveTypeToString = "NOROOTDIR" '=removed or drive has no filesystem
  Case "2"
    DriveTypeToString = "REMOVABLE"
  Case "3"
    DriveTypeToString = "FIXED"
  Case "4"
    DriveTypeToString = "REMOTE"
  Case "5"
    DriveTypeToString = "CDROM"
  Case "6"
    DriveTypeToString = "RAMDISK"
  Case Else
    DriveTypeToString = "UNKNOWN"
  End Select
End Function
