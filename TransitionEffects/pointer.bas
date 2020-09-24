Attribute VB_Name = "Module1"
Option Explicit

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'The user defined type 'SAFEARRAY2D' is used by Visual Basic for internal management multiple dimension arrays.
'The user defined type 'BITMAP' will keep some information about our picture.
'VarPtrArray' returns the memory address of an array.
'CopyMemory' copies blocks in memory from one position to another (extremely fast).
'GetObjectAPI' returns information about our bitmap, that will be written into our user defined type 'BITMAP'.


