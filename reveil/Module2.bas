Attribute VB_Name = "Module2"
 Option Explicit
 
'lpszLongPath = The complete long path and filename to convert.
'lpszShortPath = Receives the 8.3 form of the filename, terminated by a null character. This string must already be sufficiently large to receive the 8.3 filename.
'cchBuffer= The length of the string passed as lpszShortPath.
 
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
   
Public Function GetShortName(ByVal sLongFileName As String) As String
    
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
    
End Function





