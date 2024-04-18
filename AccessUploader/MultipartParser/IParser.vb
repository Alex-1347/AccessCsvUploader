Imports System.IO

<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("61efd052-27d8-9956-4297-4e4815adba52")>
<Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)>
Public Interface IParser
    Function SaveAll(path As String) As Boolean
    Function ParseMultipartFormData(content As String) As String
    Function ParseMultipartFormDataS(stream As MemoryStream) As String
    Function ParseMultipartFormDataA(data As Byte()) As String
    Property Success As Boolean
    Property ContentType As String
    Property Filename As String
    Property FileContents As Byte()
End Interface
