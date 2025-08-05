Imports System
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Namespace ComInterfaces

    <Guid("C083EDD0-F947-45D1-A96D-225C6BEF9658")>
    <ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
    Public Interface Interface1


        <DispId(200), Description("")>
        Sub Sub1(argument As String)


    End Interface

End Namespace
