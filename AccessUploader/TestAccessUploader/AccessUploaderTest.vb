Imports System.Text
Imports AccessUploader
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class AccessUploaderTest

    <TestMethod()> Public Sub ClassLoadTest1()
        Dim X As New Loader
        X.LoadCsv("E:\Projects\AccessUploader\jaypaulss-attachments\test.csv", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Projects\AccessUploader\jaypaulss-attachments\tempdb.accdb;", """", "'", ",", """", "E:\Projects\AccessUploader\jaypaulss-attachments")
    End Sub

    <TestMethod()> Public Sub ClassDeleteTest1()
        Dim X As New Loader
        X.DeleteRecord("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Projects\AccessUploader\jaypaulss-attachments\tempdb.accdb;", 1, 10)
    End Sub

    <TestMethod()> Public Sub ComTest1()
        Dim LoaderComObjType = Type.GetTypeFromProgID("Unicorn.AccessUploader")
        Dim Loader As AccessUploader.Loader = Activator.CreateInstance(LoaderComObjType)
        Loader.LoadCsv("E:\Projects\AccessUploader\jaypaulss-attachments\test.csv", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Projects\AccessUploader\jaypaulss-attachments\tempdb.accdb;", """", "'", ",", """", "E:\Projects\AccessUploader\jaypaulss-attachments")
    End Sub
    'LoadCsv(CsvFile As String, AccessConnectionString As String, Optional SpecialCharReplaceFrom As String = """", Optional SpecialCharReplaceTo As String = "'", Optional FieldDelimiter As String = ",", Optional TextDelimiter As String = """", Optional ErrorLogDirectory As String = "", Optional SkipHeaderRecords As Integer = 0, Optional StartID As Integer = 1, Optional TableName As String = "member", Optional FieldName1 As String = "Forename", Optional FieldName2 As String = "Address")
End Class