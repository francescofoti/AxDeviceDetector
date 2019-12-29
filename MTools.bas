Attribute VB_Name = "MTools"
Option Explicit

'This function is from the book "Hardcore Visual Basic 5"
'from "Sensei" Bruce McKinney. I bought the book a long time ago.
'I don't know if the code of this book is public domain or covered
'by a specific license;
'Anyway, the book's code is all over the Internet since long ago now.

'Slight own adaptation making the array static to the function.
Public Function Power2(ByVal i As Integer) As Long
  'Debug.Assert i >= 0 And i <= 31
  Static aPower2(0 To 31) As Long
  
  If aPower2(0) = 0 Then
    aPower2(0) = &H1&
    aPower2(1) = &H2&
    aPower2(2) = &H4&
    aPower2(3) = &H8&
    aPower2(4) = &H10&
    aPower2(5) = &H20&
    aPower2(6) = &H40&
    aPower2(7) = &H80&
    aPower2(8) = &H100&
    aPower2(9) = &H200&
    aPower2(10) = &H400&
    aPower2(11) = &H800&
    aPower2(12) = &H1000&
    aPower2(13) = &H2000&
    aPower2(14) = &H4000&
    aPower2(15) = &H8000&
    aPower2(16) = &H10000
    aPower2(17) = &H20000
    aPower2(18) = &H40000
    aPower2(19) = &H80000
    aPower2(20) = &H100000
    aPower2(21) = &H200000
    aPower2(22) = &H400000
    aPower2(23) = &H800000
    aPower2(24) = &H1000000
    aPower2(25) = &H2000000
    aPower2(26) = &H4000000
    aPower2(27) = &H8000000
    aPower2(28) = &H10000000
    aPower2(29) = &H20000000
    aPower2(30) = &H40000000
    aPower2(31) = &H80000000
  End If
  Power2 = aPower2(i)
End Function

