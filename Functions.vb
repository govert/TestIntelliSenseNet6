Imports ExcelDna.Integration

Public Module Functions

    <ExcelFunction(Description:="My Test Function")>
    Public Function SayHello(name As String) As Object
        Return "Hello " & name
    End Function

End Module
