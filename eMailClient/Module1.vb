Imports System.Runtime.InteropServices

Module Module1

    <DllImport("user32.dll")>
    Private Sub LockWorkStation()
    End Sub

    ''' <summary>
    ''' Call LockWorkStation here or perhaps
    ''' in a button click event
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DemoLock()
        LockWorkStation()
    End Sub

End Module