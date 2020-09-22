Attribute VB_Name = "modGeneral"
Option Explicit

' /*
' ** Allocates the specified number of bytes from the heap.
' */
Public Declare Function GlobalAlloc _
    Lib "kernel32" ( _
        ByVal wFlags As Long, _
        ByVal dwBytes As Long) As Long

' /*
' ** Locks a global memory object and returns a pointer to
' ** the first byte of the bject's memory block.
' ** The memory block associated with a locked object cannot
' ** be moved or discarded.
'*/
Public Declare Function GlobalLock _
    Lib "kernel32" ( _
        ByVal hmem As Long) As Long

' /*
' ** Frees the specificed global memory object and
' ** invalidates its handle
' */
Public Declare Function GlobalFree _
    Lib "kernel32" ( _
        ByVal hmem As Long) As Long

Public Declare Sub CopyPtrFromStruct _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ByVal ptr As Long, _
        struct As Any, _
        ByVal cb As Long)
        

Public Declare Sub CopyStructFromPtr _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        struct As Any, _
        ByVal ptr As Long, _
        ByVal cb As Long)
        
Public Declare Sub memcpy _
    Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        ptr As Any, _
        ptr2 As Any, _
        ByVal cb As Long)
        
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long

Public Function StripNulls(STR) As String
  Dim i As Long
  For i = 0 To UBound(STR)
    If STR(i) <> 0 Then StripNulls = StripNulls & Chr(STR(i))
  Next i
End Function

