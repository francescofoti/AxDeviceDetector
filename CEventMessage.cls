VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This is a pure data class, used to push and pop messages
'to/from a CEventMessage queue.
'It mimics the TEventMessage structure (or type definition).

Public EventID      As Integer
Public DeviceType   As String
Public DriveCt      As Integer
Public DriveLetters As String
Public DriveTypes   As String

