VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeviceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If Win64 Then
Private Declare PtrSafe Function DICreateDeviceInfo Lib "deviceinfo64.dll" () As LongPtr
Private Declare PtrSafe Sub DIDestroyDeviceInfo Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr)
Private Declare PtrSafe Function DIQueryDeviceInfo Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr, ByVal psDrivePath As LongPtr) As Integer
Private Declare PtrSafe Function DIGetVersion Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Long
Private Declare PtrSafe Function DIGetDeviceType Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Byte
Private Declare PtrSafe Function DIGetDeviceTypeModifier Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Byte
Private Declare PtrSafe Function DIGetRemovableMedia Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Byte
Private Declare PtrSafe Function DIGetCommandQueueing Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Byte
Private Declare PtrSafe Function DIGetBusType Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Byte
Private Declare PtrSafe Function DIGetLastError Lib "deviceinfo64.dll" (ByVal phDevInfo As LongPtr) As Long
Private Declare PtrSafe Function DIGetLastErrorText Lib "deviceinfo64.dll" ( _
  ByVal phDevInfo As LongPtr, _
  ByVal plRetOutputBuf As LongPtr, _
  ByVal piOutputBufLen As Integer) As Long
Private Declare PtrSafe Sub DIGetVendorID Lib "deviceinfo64.dll" ( _
  ByVal phDevInfo As LongPtr, _
  ByVal plRetOutputBuf As LongPtr, _
  ByVal piOutputBufLen As Integer)
Private Declare PtrSafe Sub DIGetProductID Lib "deviceinfo64.dll" ( _
  ByVal phDevInfo As LongPtr, _
  ByVal plRetOutputBuf As LongPtr, _
  ByVal piOutputBufLen As Integer)
Private Declare PtrSafe Sub DIGetProductRevision Lib "deviceinfo64.dll" ( _
  ByVal phDevInfo As LongPtr, _
  ByVal plRetOutputBuf As LongPtr, _
  ByVal piOutputBufLen As Integer)
Private Declare PtrSafe Sub DIGetSerialNumber Lib "deviceinfo64.dll" ( _
  ByVal phDevInfo As LongPtr, _
  ByVal plRetOutputBuf As LongPtr, _
  ByVal piOutputBufLen As Integer)
#Else
Private Declare Function DICreateDeviceInfo Lib "deviceinfo.dll" () As Long
Private Declare Sub DIDestroyDeviceInfo Lib "deviceinfo.dll" (ByVal phDevInfo As Long)
Private Declare Function DIQueryDeviceInfo Lib "deviceinfo.dll" (ByVal phDevInfo As Long, ByVal psDrivePath As Long) As Integer
Private Declare Function DIGetVersion Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Long
Private Declare Function DIGetDeviceType Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Byte
Private Declare Function DIGetDeviceTypeModifier Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Byte
Private Declare Function DIGetRemovableMedia Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Byte
Private Declare Function DIGetCommandQueueing Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Byte
Private Declare Function DIGetBusType Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Byte
Private Declare Function DIGetLastError Lib "deviceinfo.dll" (ByVal phDevInfo As Long) As Long
Private Declare Function DIGetLastErrorText Lib "deviceinfo.dll" ( _
  ByVal phDevInfo As Long, _
  ByVal plRetOutputBuf As Long, _
  ByVal piOutputBufLen As Integer) As Long
Private Declare Sub DIGetVendorID Lib "deviceinfo.dll" ( _
  ByVal phDevInfo As Long, _
  ByVal plRetOutputBuf As Long, _
  ByVal piOutputBufLen As Integer)
Private Declare Sub DIGetProductID Lib "deviceinfo.dll" ( _
  ByVal phDevInfo As Long, _
  ByVal plRetOutputBuf As Long, _
  ByVal piOutputBufLen As Integer)
Private Declare Sub DIGetProductRevision Lib "deviceinfo.dll" ( _
  ByVal phDevInfo As Long, _
  ByVal plRetOutputBuf As Long, _
  ByVal piOutputBufLen As Integer)
Private Declare Sub DIGetSerialNumber Lib "deviceinfo.dll" ( _
  ByVal phDevInfo As Long, _
  ByVal plRetOutputBuf As Long, _
  ByVal piOutputBufLen As Integer)
#End If

Private Const MAX_ERRSTR_LENGTH     As Integer = 512
Private Const INFOSTRING_MAX_LENGTH As Integer = 128

Private mlVersion            As Long
Private mbDeviceType         As Byte
Private mbDeviceTypeModifier As Byte
Private mbRemovableMedia     As Byte
Private mbCommandQueueing    As Byte
Private mbBusType            As Byte
Private msVendorID           As String
Private msProductID          As String
Private msProductRevision    As String
Private msSerialNumber       As String

' Class error context
Private mlErr     As Long
Private msErr     As String
Private msErrCtx  As String
 
Private Sub ClearErr()
  mlErr = 0&
  msErr = ""
  msErrCtx = ""
End Sub
 
Private Sub SetErr(ByVal psErrCtx As String, ByVal plErr As Long, ByVal psErr As String)
  mlErr = plErr
  msErr = psErr
  msErrCtx = psErrCtx
End Sub
 
Public Property Get LastErr() As Long
  LastErr = mlErr
End Property
 
Public Property Get LastErrDesc() As String
  LastErrDesc = msErr
End Property
 
Public Property Get LastErrCtx() As String
  LastErrCtx = msErrCtx
End Property

'All read only properties
Public Property Get Version() As Long: Version = mlVersion: End Property
Public Property Get DeviceType() As Byte: DeviceType = mbDeviceType: End Property
Public Property Get DeviceTypeModifier() As Byte: DeviceTypeModifier = mbDeviceTypeModifier: End Property
Public Property Get RemovableMedia() As Byte: RemovableMedia = mbRemovableMedia: End Property
Public Property Get CommandQueueing() As Byte: CommandQueueing = mbCommandQueueing: End Property
Public Property Get BusType() As Byte: BusType = mbBusType: End Property
Public Property Get VendorID() As String: VendorID = msVendorID: End Property
Public Property Get ProductID() As String: ProductID = msProductID: End Property
Public Property Get ProductRevision() As String: ProductRevision = msProductRevision: End Property
Public Property Get SerialNumber() As String: SerialNumber = msSerialNumber: End Property

Private Function MakeBufferString(ByVal piBufferSize As Integer) As String
  MakeBufferString = Space$(piBufferSize)
End Function

'Cut string before trailing chr$(0)
Private Function CtoVB(ByRef pszString As String) As String
  Dim i   As Long
  i = InStr(pszString, Chr$(0))
  If i Then
    CtoVB = Left$(pszString, i - 1&)
  Else
    CtoVB = pszString
  End If
End Function

Public Function GetDeviceInformation(ByVal psDriveLetter As String) As Boolean
  Const LOCAL_ERR_CTX As String = "GetDeviceInformation"
  Const ERR_MEMORY As Long = -1&
  
  On Error GoTo GetDeviceInformation_Err
  ClearErr
  
  Dim fOK             As Boolean
  Dim sDrivePath      As String
  Dim sErrString      As String
  Dim lErrCode        As Long
  Dim sBuffer         As String
  #If Win64 Then
  Dim hDevInfo        As LongPtr
  #Else
  Dim hDevInfo        As Long
  #End If
  
  sDrivePath = "\\.\" & psDriveLetter & ":" & ChrW$(0)
  hDevInfo = DICreateDeviceInfo()
  If hDevInfo = 0 Then
    SetErr LOCAL_ERR_CTX, ERR_MEMORY, "Couldn't allocate memory to query for device information"
    Exit Function
  End If
  
  fOK = DIQueryDeviceInfo(hDevInfo, StrPtr(sDrivePath))
  If Not fOK Then
    sErrString = MakeBufferString(MAX_ERRSTR_LENGTH)
    lErrCode = DIGetLastErrorText(hDevInfo, StrPtr(sErrString), MAX_ERRSTR_LENGTH)
    SetErr LOCAL_ERR_CTX, lErrCode, Trim$(CtoVB(sErrString))
    GoTo GetDeviceInformation_Exit
  End If
  
  mlVersion = DIGetVersion(hDevInfo)
  mbDeviceType = DIGetDeviceType(hDevInfo)
  mbDeviceTypeModifier = DIGetDeviceTypeModifier(hDevInfo)
  mbRemovableMedia = DIGetRemovableMedia(hDevInfo)
  mbCommandQueueing = DIGetCommandQueueing(hDevInfo)
  mbBusType = DIGetBusType(hDevInfo)
  'retrieve strings
  sBuffer = MakeBufferString(INFOSTRING_MAX_LENGTH)
  DIGetVendorID hDevInfo, StrPtr(sBuffer), INFOSTRING_MAX_LENGTH
  msVendorID = Trim$(CtoVB(sBuffer))
  sBuffer = MakeBufferString(INFOSTRING_MAX_LENGTH)
  DIGetProductID hDevInfo, StrPtr(sBuffer), INFOSTRING_MAX_LENGTH
  msProductID = Trim$(CtoVB(sBuffer))
  sBuffer = MakeBufferString(INFOSTRING_MAX_LENGTH)
  DIGetProductRevision hDevInfo, StrPtr(sBuffer), INFOSTRING_MAX_LENGTH
  msProductRevision = Trim$(CtoVB(sBuffer))
  sBuffer = MakeBufferString(INFOSTRING_MAX_LENGTH)
  DIGetSerialNumber hDevInfo, StrPtr(sBuffer), INFOSTRING_MAX_LENGTH
  msSerialNumber = Trim$(CtoVB(sBuffer))
  
  GetDeviceInformation = True
  
GetDeviceInformation_Exit:
  If hDevInfo Then
    DIDestroyDeviceInfo hDevInfo
  End If
  Exit Function

GetDeviceInformation_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume GetDeviceInformation_Exit
End Function

