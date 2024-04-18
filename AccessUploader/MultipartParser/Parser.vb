Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("8d5e44c5-1e0a-a2a7-4d5b-dd7f0827cdcf")>
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.AutoDispatch)>
<Runtime.InteropServices.ProgId("Unicorn.MultipartParser")>
Public Class Parser
    Implements IParser
    Public Function SaveAll(path As String) As Boolean Implements IParser.SaveAll
        Try
            My.Computer.FileSystem.WriteAllBytes(path, m_FileContents, False)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function ParseMultipartFormDataS(stream As MemoryStream) As String Implements IParser.ParseMultipartFormDataS

        ' Read the stream into a byte array
        Dim data As Byte() = ToByteArray(stream)
        ' Copy to a string for header parsing
        Return ParseMultipartFormDataA(data)

    End Function

    Public Function ParseMultipartFormDataA(data As Byte()) As String Implements IParser.ParseMultipartFormDataA

        Dim encoding As Encoding = Encoding.UTF8
        Dim content As String = encoding.GetString(data)
        Return ParseMultipartFormData(content)

    End Function

    Public Function ParseMultipartFormData(content As String) As String Implements IParser.ParseMultipartFormData
        Me.Success = False
        Dim encoding As Encoding = Encoding.UTF8
        Dim Data As Byte() = Encoding.UTF8.GetBytes(content)
        Try
            ' The first line should contain the delimiter
            Dim delimiterEndIndex As Integer = content.IndexOf(vbCr & vbLf)
            If delimiterEndIndex > -1 Then
                Dim delimiter As String = content.Substring(0, content.IndexOf(vbCr & vbLf))
                ' Look for Content-Type
                Dim re As New Regex("(?<=Content\-Type:)(.*?)(?=\r\n\r\n)")
                Dim contentTypeMatch As Match = re.Match(content)
                ' Look for filename
                re = New Regex("(?<=filename\=\"")(.*?)(?=\"")")
                Dim filenameMatch As Match = re.Match(content)
                ' Did we find the required values?
                If contentTypeMatch.Success AndAlso filenameMatch.Success Then
                    ' Set properties
                    Me.ContentType = contentTypeMatch.Value.Trim()
                    Me.Filename = filenameMatch.Value.Trim()
                    ' Get the start & end indexes of the file contents
                    Dim startIndex As Integer = contentTypeMatch.Index + contentTypeMatch.Length + (vbCr & vbLf & vbCr & vbLf).Length
                    Dim delimiterBytes As Byte() = encoding.GetBytes(vbCr & vbLf & delimiter)
                    Dim endIndex As Integer = IndexOf(Data, delimiterBytes, startIndex)
                    Dim contentLength As Integer = endIndex - startIndex
                    ' Extract the file contents from the byte array
                    Dim fileData As Byte() = New Byte(contentLength - 1) {}
                    Buffer.BlockCopy(Data, startIndex, fileData, 0, contentLength)
                    Me.FileContents = fileData
                    Me.Success = True
                End If
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function IndexOf(searchWithin As Byte(), serachFor As Byte(), startIndex As Integer) As Integer
        Dim index As Integer = 0
        Dim startPos As Integer = Array.IndexOf(searchWithin, serachFor(0), startIndex)
        If startPos <> -1 Then
            While (startPos + index) < searchWithin.Length
                If searchWithin(startPos + index) = serachFor(index) Then
                    index += 1
                    If index = serachFor.Length Then
                        Return startPos
                    End If
                Else
                    startPos = Array.IndexOf(Of Byte)(searchWithin, serachFor(0), startPos + index)
                    If startPos = -1 Then
                        Return -1
                    End If
                    index = 0
                End If
            End While
        End If
        Return -1
    End Function
    Private Function ToByteArray(stream As Stream) As Byte()
        Dim buffer As Byte() = New Byte(32767) {}
        Using ms As New MemoryStream()
            While True
                Dim read As Integer = stream.Read(buffer, 0, buffer.Length)
                If read <= 0 Then
                    Return ms.ToArray()
                End If
                ms.Write(buffer, 0, read)
            End While
        End Using
    End Function
    Public Property Success As Boolean Implements IParser.Success
        Get
            Return m_Success
        End Get
        Private Set(value As Boolean)
            m_Success = value
        End Set
    End Property
    Private m_Success As Boolean
    Public Property ContentType As String Implements IParser.ContentType
        Get
            Return m_ContentType
        End Get
        Private Set(value As String)
            m_ContentType = value
        End Set
    End Property
    Private m_ContentType As String
    Public Property Filename As String Implements IParser.Filename
        Get
            Return m_Filename
        End Get
        Private Set(value As String)
            m_Filename = value
        End Set
    End Property
    Private m_Filename As String
    Public Property FileContents As Byte() Implements IParser.FileContents
        Get
            Return m_FileContents
        End Get
        Private Set(value As Byte())
            m_FileContents = value
        End Set
    End Property
    Private m_FileContents As Byte()


End Class

