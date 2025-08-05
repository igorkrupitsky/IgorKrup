Imports System.Runtime.InteropServices

Module KeyboardInjector
    <DllImport("Keyboard.dll", CallingConvention:=CallingConvention.Cdecl)>
    Public Sub SendKeyScan(scanCode As UShort, isDown As Boolean, extended As Boolean)
    End Sub
End Module