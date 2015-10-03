Attribute VB_Name = "modMemory"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Memory(65536) As Byte

Public Sub Initialise()
    Memory(131072 \ 8) = 255
    Memory(131080 \ 8) = 129
    Memory(Temp() \ 8) = &H0
End Sub
