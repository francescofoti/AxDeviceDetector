VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CEventQueue Class
'
'Responsibility
'The responsibility of this class is to store a set of event messages
' to then give them back in the order they arrive (FIFO).
'
'How
'The data of an event message is stored internally by the queue in
'structured variables of type TEventMessage (defined in a module outside this class).
'Unfortunately, VIsual Basic (5 here) doesn't support passing structures (ie Types)
'in function parameters.
'So, to push and pop elements, we use transient objects of CEventMessage class,
'which is a class that mimics the TEventMessage type structure, and we copy
'the data between the class and the structure when needed.
'This way, we do not have to juggle with dynamic memory and object references.
'
'This is an adaptation of the class module presented in my blog article here:
' https://francescofoti.com/2019/12/how-to-implement-a-circular-queue-in-vba-excel/
'where you can also see an animated explanation of the class inner workings.

Private Type TQHeader
  iQCount       As Integer
  iQBack        As Integer
  iQFront       As Integer
  iMaxCapacity  As Integer
End Type

'The queue header holds runtime information about the queue
Private mTHeader    As TQHeader

'The internal data array holds the queue items
Private matItem()   As TEventMessage

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
  'nothing to do for now
End Sub

Private Sub Class_Terminate()
  'nothing to do for now
End Sub

' Members

Public Property Get Count() As Integer
  Count = mTHeader.iQCount
End Property

Public Property Get MaxCapacity() As Integer
  MaxCapacity = mTHeader.iMaxCapacity
End Property

'A queue is created with the given maximum capacity that cannot be changed after creation.
Public Function CreateQueue(ByVal piMaxCapacity As Integer) As Boolean
  Const LOCAL_ERR_CTX As String = "CreateQueue"
  On Error GoTo CreateQueue_Err
  ClearErr
  
  If piMaxCapacity < 1 Then
    SetErr LOCAL_ERR_CTX, -1&, "A minimum capacity of 1 is required to create the queue"
    Exit Function
  End If
  
  ReDim matItem(0 To piMaxCapacity - 1) As TEventMessage
  
  With mTHeader
    .iMaxCapacity = piMaxCapacity
    .iQBack = 0
    .iQCount = 0
    .iQFront = 0
  End With
  
  CreateQueue = True
  
CreateQueue_Exit:
  Exit Function

CreateQueue_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume CreateQueue_Exit
End Function

Public Sub Clear()
  On Error Resume Next
  Dim i   As Integer
  Dim tEmptyMessage As TEventMessage
  
  With mTHeader
    .iQBack = 0
    .iQCount = 0
    .iQFront = 0
    'clear internal array
    For i = 0 To .iMaxCapacity - 1
      matItem(i) = tEmptyMessage
    Next i
  End With
End Sub

'Push an item in the queue.
'Returns True if the item is in the queue.
Public Function QPush(ByRef poItem As CEventMessage) As Boolean
  Dim tMessage As TEventMessage
  ClearErr
  
  If mTHeader.iMaxCapacity = mTHeader.iQCount Then
    SetErr "QPush", -1&, "Queue is full"
    Exit Function
  End If
  
  'object to structure
  With tMessage
    .iDriveCt = poItem.DriveCt
    .iEventID = poItem.EventID
    .sDeviceType = poItem.DeviceType
    .sDriveLetters = poItem.DriveLetters
    .sDriveTypes = poItem.DriveTypes
  End With
  
  If mTHeader.iQCount = 0 Then
    With mTHeader
      .iQBack = 1
      .iQFront = 1
      .iQCount = 0
    End With
    matItem(0) = tMessage
  Else
    With mTHeader
      
      If .iQCount = .iMaxCapacity Then
        SetErr "QPush", -1, "Queue is full"
        Exit Function
      End If
      
      If .iQBack < .iMaxCapacity Then
        .iQBack = .iQBack + 1
      Else
        .iQBack = 1
      End If
      matItem(.iQBack - 1) = tMessage
      
    End With
  End If
  mTHeader.iQCount = mTHeader.iQCount + 1
  
  QPush = True
End Function

'Returns True if a queued item is popped out of the queue.
'The popped out item is copied in poRetItem.
Public Function QPop(ByRef poRetItem As CEventMessage) As Boolean
  Dim tMessage      As TEventMessage
  Dim tEmptyMessage As TEventMessage
  ClearErr
  
  If mTHeader.iQCount = 0 Then
    SetErr "QPop", -1&, "Queue is empty"
    Exit Function
  End If
  
  tMessage = matItem(mTHeader.iQFront - 1)
  matItem(mTHeader.iQFront - 1) = tEmptyMessage
  
  With mTHeader
    If .iQCount > 0 Then
      .iQFront = .iQFront + 1
      If .iQFront > .iMaxCapacity Then
        .iQFront = 1
      End If
      .iQCount = .iQCount - 1
    Else
      Me.Clear
    End If
  End With
  
  'structure to object
  With tMessage
    poRetItem.DriveCt = .iDriveCt
    poRetItem.EventID = .iEventID
    poRetItem.DeviceType = .sDeviceType
    poRetItem.DriveLetters = .sDriveLetters
    poRetItem.DriveTypes = .sDriveTypes
  End With
  
  QPop = True
End Function

