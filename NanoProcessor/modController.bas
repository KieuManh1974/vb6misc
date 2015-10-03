Attribute VB_Name = "modController"
Option Explicit

Public Sub Main()
    modMemory.Initialise
    modNano.Initialise
    modCompile.Compile
    
    InitialiseCounters
    
    StartCounter
    
    modNano.Execute
End Sub
