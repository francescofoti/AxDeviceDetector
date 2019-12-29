VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implement this interface to handle session notification messages in your form
Function WindowProc(ByVal plhWnd As Long, _
                    ByVal plMsg As Long, _
                    ByVal pwParam As Long, _
                    ByVal plParam As Long) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
End Function

