<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("61e27d81-fd05-9956-4297-4e4815adba52")>
<Runtime.InteropServices.InterfaceType(Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)>
Public Interface ILoader
    Function LoadCsv(CsvFile As String, AccessConnectionString As String, Optional SpecialCharReplaceFrom As String = """", Optional SpecialCharReplaceTo As String = "'", Optional FieldDelimiter As String = ",", Optional TextDelimiter As String = """", Optional ErrorLogDirectory As String = "", Optional SkipHeaderRecords As Integer = 0, Optional StartID As Integer = 1, Optional TableName As String = "member", Optional FieldName1 As String = "Forename", Optional FieldName2 As String = "Address") As Integer
    Sub DeleteRecord(AccessConnectionString As String, FromIDInclude As Integer, ToIDExclude As Integer, Optional ErrorLogDirectory As String = "", Optional TableName As String = "member")
End Interface
