Imports System.IO
Imports System.Security.Cryptography.X509Certificates

<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("1e0a44c5-8d5e-a2a7-4d5b-dd7f0827cdcf")>
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.AutoDispatch)>
<Runtime.InteropServices.ProgId("Unicorn.AccessUploader")>
Public Class Loader
    Implements ILoader

    Public Function LoadCsv(CsvFile As String, AccessConnectionString As String, Optional SpecialCharReplaceFrom As String = """", Optional SpecialCharReplaceTo As String = "'", Optional FieldDelimiter As String = ",", Optional TextDelimiter As String = """", Optional ErrorLogDirectory As String = "", Optional SkipHeaderRecords As Integer = 0, Optional StartID As Integer = 1, Optional TableName As String = "member", Optional FieldName1 As String = "Forename", Optional FieldName2 As String = "Address") As Integer Implements ILoader.LoadCsv
        Dim TempFileName = IO.Path.Combine(ErrorLogDirectory, Now.Ticks & ".txt")
        Dim Data As IEnumerable(Of OneLine)
        Try
            Data = From Lines In File.ReadLines(CsvFile)
                   Let X = Lines.Split({FieldDelimiter}, StringSplitOptions.RemoveEmptyEntries).Skip(SkipHeaderRecords)
                   Select New OneLine With {.Forename = X(0), .Address = X(1)}

            If ErrorLogDirectory <> "" Then
                My.Computer.FileSystem.WriteAllText(TempFileName, $"{Now} Read {Data.ToList.Count} records from {CsvFile}{vbCrLf}", True)
            End If
            Dim Cn1 As New OleDb.OleDbConnection(AccessConnectionString)
            Cn1.Open()
            Dim Cmd As OleDb.OleDbCommand
            Data.ToList.ForEach(Sub(X As OneLine)
                                    'Debug.Print($"{X.Forename}{FieldDelimiter}{X.Address}")
                                    Cmd = New OleDb.OleDbCommand($"INSERT INTO {TableName} (ID, {FieldName1}, {FieldName2}) VALUES ('{StartID}', {Replacement(X.Forename, SpecialCharReplaceFrom, SpecialCharReplaceTo)},{Replacement(X.Address, SpecialCharReplaceFrom, SpecialCharReplaceTo)});", Cn1)
                                    If ErrorLogDirectory <> "" Then
                                        My.Computer.FileSystem.WriteAllText(TempFileName, Cmd.CommandText & vbCrLf, True)
                                    Else
                                        Debug.Print(Cmd.CommandText)
                                    End If
                                    Cmd.ExecuteNonQuery()
                                    Cmd.Dispose()
                                    StartID = StartID + 1
                                End Sub)
            Cn1.Close()
            Cn1.Dispose()
        Catch ex As Exception
            If ErrorLogDirectory <> "" Then
                My.Computer.FileSystem.WriteAllText(TempFileName, ex.Message, True)
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
        Return Data?.Count
    End Function

    Function Replacement(Field As String, Optional From As String = "", Optional [To] As String = "") As String
        If From <> "" And [To] <> "" Then Return Field.Replace(From, [To]) Else Return Field
    End Function

    Public Sub DeleteRecord(AccessConnectionString As String, FromIDInclude As Integer, ToIDExclude As Integer, Optional ErrorLogDirectory As String = "", Optional TableName As String = "member") Implements ILoader.DeleteRecord
        Dim TempFileName = IO.Path.Combine(ErrorLogDirectory, Now.Ticks & ".txt")
        Try
            Dim Cn1 As New OleDb.OleDbConnection(AccessConnectionString)
            Cn1.Open()
            Dim Cmd As New OleDb.OleDbCommand($"DELETE FROM {TableName} WHERE ID>={FromIDInclude} AND ID<{ToIDExclude};", Cn1)
            Cmd.ExecuteNonQuery()
            Cn1.Close()
            Cn1.Dispose()
        Catch ex As Exception
            If ErrorLogDirectory <> "" Then
                My.Computer.FileSystem.WriteAllText(TempFileName, ex.Message, True)
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
    End Sub
End Class

Public Class OneLine
    Property Forename As String
    Property Address As String
End Class
