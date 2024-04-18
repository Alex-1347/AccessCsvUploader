Imports System.IO
Imports System.Text
Imports AccessUploader
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports MultipartParser

<TestClass()> Public Class MultipartParserTest

    <TestMethod()> Public Sub ClassParserTestS()
        Dim Request As String = My.Computer.FileSystem.ReadAllText(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, "Request.txt"))
        Dim X As New Parser
        X.ParseMultipartFormDataS(New MemoryStream(Encoding.UTF8.GetBytes(Request)))

        X.SaveAll(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, Now.Ticks & ".txt"))
    End Sub

    <TestMethod()> Public Sub ComTestS()
        Dim LoaderComObjType = Type.GetTypeFromProgID("Unicorn.MultipartParser")
        Dim Loader As MultipartParser.Parser = Activator.CreateInstance(LoaderComObjType)
        Dim Request As String = My.Computer.FileSystem.ReadAllText(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, "Request.txt"))
        Dim X As New Parser
        X.ParseMultipartFormDataS(New MemoryStream(Encoding.UTF8.GetBytes(Request)))
        X.SaveAll(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, Now.Ticks & ".txt"))
    End Sub

    <TestMethod()> Public Sub ClassParserTestA()
        Dim Request As String = My.Computer.FileSystem.ReadAllText(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, "Request.txt"))
        Dim X As New Parser
        X.ParseMultipartFormDataA(Encoding.UTF8.GetBytes(Request))
        X.SaveAll(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, Now.Ticks & ".txt"))
    End Sub

    <TestMethod()> Public Sub ComTestA()
        Dim LoaderComObjType = Type.GetTypeFromProgID("Unicorn.MultipartParser")
        Dim Loader As MultipartParser.Parser = Activator.CreateInstance(LoaderComObjType)
        Dim Request As String = My.Computer.FileSystem.ReadAllText(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, "Request.txt"))
        Dim X As New Parser
        X.ParseMultipartFormDataA(Encoding.UTF8.GetBytes(Request))
        X.SaveAll(IO.Path.Combine(My.Computer.FileSystem.CurrentDirectory, Now.Ticks & ".txt"))
    End Sub


End Class