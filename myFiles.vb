Class myFiles

#Region " Compress File "

    Public Shared Function Compress(ByVal source As String, ByVal destination As String) As Boolean
        If Delete(destination, False) = False Then Return False

        ' Create the streams and byte arrays needed
        Dim buffer As Byte() = Nothing
        Dim sourceStream As IO.FileStream = Nothing
        Dim destinationStream As IO.FileStream = Nothing
        Dim compressedStream As IO.Compression.GZipStream = Nothing
        Try
            ' Read the bytes from the source file into a byte array
            sourceStream = New IO.FileStream(source, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
            ' Read the source stream values into the buffer
            buffer = New Byte(CInt(sourceStream.Length)) {}
            Dim checkCounter As Integer = sourceStream.Read(buffer, 0, buffer.Length)
            ' Open the FileStream to write to
            destinationStream = New IO.FileStream(destination, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
            ' Create a compression stream pointing to the destiantion stream
            compressedStream = New IO.Compression.GZipStream(destinationStream, IO.Compression.CompressionMode.Compress, True)
            'Now write the compressed data to the destination file
            compressedStream.Write(buffer, 0, buffer.Length)
        Catch ex As ApplicationException
            MessageBox.Show(ex.Message, "ZIP file compressing")
            Return False
        Finally
            ' Make sure we allways close all streams   
            If Not (sourceStream Is Nothing) Then
                sourceStream.Close()
            End If
            If Not (compressedStream Is Nothing) Then
                compressedStream.Close()
            End If
            If Not (destinationStream Is Nothing) Then
                destinationStream.Close()
            End If
        End Try
        Return True
    End Function
#End Region

#Region " Decompress File or Embedded File "

    Public Shared Function Decompress(ByVal source As String, ByVal destination As String) As Boolean
        If Delete(destination, False) = False Then Return False
        Dim sourceStream As IO.Stream
        If Exist(source) Then
            sourceStream = New IO.FileStream(source, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
        Else
            sourceStream = Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("RootSpace." & source & ".zip")
        End If

        Dim Reader As New System.IO.BinaryReader(sourceStream)
        Dim fileByte() As Byte = Reader.ReadBytes(CInt(sourceStream.Length))
        ' Create the streams and byte arrays needed
        Dim memStream As IO.MemoryStream = New IO.MemoryStream(fileByte)
        Dim destinationStream As IO.FileStream = Nothing
        Dim decompressedStream As IO.Compression.GZipStream = Nothing
        Dim quartetBuffer As Byte() = Nothing
        Try
            ' Read in the compressed source stream
            'sourceStream = New FileStream(sourceFile, FileMode.Open)
            ' Create a compression stream pointing to the destiantion stream
            decompressedStream = New IO.Compression.GZipStream(memStream, IO.Compression.CompressionMode.Decompress, True)
            ' Read the footer to determine the length of the destination file
            quartetBuffer = New Byte(4) {}
            Dim position As Integer = CType(memStream.Length, Integer) - 4
            memStream.Position = position
            memStream.Read(quartetBuffer, 0, 4)
            memStream.Position = 0
            Dim checkLength As Integer = BitConverter.ToInt32(quartetBuffer, 0)
            Dim buffer(checkLength + 100) As Byte
            Dim offset As Integer = 0
            Dim total As Integer = 0
            ' Read the compressed data into the buffer
            While True
                Dim bytesRead As Integer = decompressedStream.Read(buffer, offset, 100)
                If bytesRead = 0 Then
                    Exit While
                End If
                offset += bytesRead
                total += bytesRead
            End While
            ' Now write everything to the destination file
            destinationStream = New IO.FileStream(destination, IO.FileMode.Create)
            destinationStream.Write(buffer, 0, total - 1)
            ' and flush everyhting to clean out the buffer
            destinationStream.Flush()
            Decompress = True
        Catch ex As ApplicationException
            Decompress = False
            MessageBox.Show("Access denied.", "File decompressing.")
        Finally
            ' Make sure we allways close all streams
            If Not (Reader Is Nothing) Then
                Reader.Close()
            End If
            If Not (sourceStream Is Nothing) Then
                sourceStream.Close()
            End If
            If Not (memStream Is Nothing) Then
                memStream.Close()
            End If
            If Not (sourceStream Is Nothing) Then
                sourceStream.Close()
            End If
            If Not (decompressedStream Is Nothing) Then
                decompressedStream.Close()
            End If
            If Not (destinationStream Is Nothing) Then
                destinationStream.Close()
            End If
        End Try
    End Function

#End Region

#Region " Safename "

    Public Shared Function GetNameSafe(Name As String) As String
        Name = Name.Replace("?", "")
        Name = Name.Replace("'", "")
        Name = Name.Replace("*", "")
        Name = Name.Replace("/", "")
        Name = Name.Replace("<", "")
        Name = Name.Replace(">", "")
        Name = Name.Replace(":", "")
        Name = Name.Replace("\", "")
        Name = Name.Replace("'", "")
        Name = Name.Replace(Chr(34), "")
        Return Name
    End Function

    Public Shared Function isNameSafe(name As String) As Boolean
        If name.Contains("?") Or name.Contains("'") Or name.Contains("*") Or name.Contains("/") Or name.Contains("<") Or name.Contains(">") Or name.Contains(":") Or name.Contains("\") Or name.Contains(Chr(34)) Then Return False
        Return True
    End Function

    Public Shared Function isPathSafe(path As String) As Boolean
        If path.Contains("?") Or path.Contains("'") Or path.Contains("*") Or path.Contains("/") Or path.Contains("<") Or path.Contains(">") Then Return False
        Return True
    End Function
    Public Shared Function isSearchSafe(name As String) As Boolean
        If name.Contains("!") Or name.Contains("'") Or name.Contains("/") Or name.Contains("<") Or name.Contains(">") Or name.Contains(":") Or name.Contains("\") Or name.Contains(Chr(34)) Then Return False
        Return True
    End Function

#End Region

#Region " Cleansename "
    Public Shared Function GetCleanseName(Name As String) As String
        Dim a, b As Integer
        a = Name.LastIndexOf("(")
        b = Name.LastIndexOf(")")
        If Not a = -1 And Not b = -1 Then
            Name = Name.Substring(0, a) & Name.Substring(b + 1, Name.Length - b - 1)
        End If
        a = Name.LastIndexOf("_")
        If Not a = -1 Then Name = Name.Substring(0, a)
        Name = Name.Replace("_", " ")
        Name = Name.Replace(".", " ")
        Return Name
    End Function

#End Region

#Region " Join "

    Private Shared Function FixPath(Folder As String) As String
        If Folder.StartsWith("\") Then Folder = Folder.Substring(1, Folder.Length - 1)
        Return Folder
    End Function

    Public Shared Function Join(Folder As String, File As String) As String
        Return System.IO.Path.Combine(FixPath(Folder), FixPath(File))
    End Function

    Public Shared Function Join(Folder1 As String, Folder2 As String, File As String) As String
        Return System.IO.Path.Combine(FixPath(Folder1), FixPath(Folder2), FixPath(File))
    End Function

#End Region

#Region " Atributes "

    Public Shared Function Size(ByVal Path As String) As Long
        Dim info As New System.IO.FileInfo(Path)
        Return info.Length
    End Function

    Public Shared Function DateCreation(ByVal Path As String) As Date
        Dim info As New System.IO.FileInfo(Path)
        Return info.CreationTime
    End Function

    Public Shared Function DateChange(ByVal Path As String) As Date
        Dim info As New System.IO.FileInfo(Path)
        Return info.LastWriteTime
    End Function

    Public Shared Function DateOpen(ByVal Path As String) As Date
        Dim info As New System.IO.FileInfo(Path)
        Return info.LastAccessTime
    End Function

    Public Shared Function DateSetOpen(Path As String, Datum As Date) As Boolean
        DateSetOpen = False
        Dim info As New System.IO.FileInfo(Path)
        If info.Exists Then
            Try
                info.LastAccessTime = Datum
                DateSetOpen = True
            Catch Err As Exception
                MessageBox.Show("File " & Path & " is inaccessible.", "LastAccessTime")
            End Try
        End If
    End Function

#End Region

#Region " Embedded Resource "

    Public Shared Function ReadEmbeddedResource(FileName As String) As Byte()
        Dim Stream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("RootSpace." & FileName)
        Dim Reader As New System.IO.BinaryReader(Stream)
        Return Reader.ReadBytes(CInt(Stream.Length))
    End Function

#End Region

    Public Shared Function ToStream(Filename As String) As IO.MemoryStream
        Dim ms As New IO.MemoryStream
        Try
            Using stream As New IO.FileStream(Filename, IO.FileMode.Open, IO.FileAccess.Read)
                stream.CopyTo(ms)
            End Using
        Catch
        End Try
        ms.Position = 0
        Return ms
    End Function

    Public Shared Function DiskType(Filename As String) As Integer
        Return New System.IO.DriveInfo(Filename.Substring(0, 1)).DriveType
    End Function

    Public Shared Function Extension(Filename As String) As String
        If Filename = "" Then Return ""
        Filename = RemoveQuotationMarks(Filename)
        Return Strings.Right(Filename, Filename.Length - Filename.LastIndexOf("."))
    End Function

    Public Shared Function Name(Filename As String, Optional WithExtension As Boolean = True) As String
        If Filename = "" Then Return ""
        Filename = RemoveQuotationMarks(Filename)
        Dim Fols() As String = Split(Filename, "\".ToCharArray)
        Dim dName As String = Fols(UBound(Fols))
        If WithExtension = False Then
            If Not dName.LastIndexOf(".".ToCharArray) = -1 Then
                dName = Strings.Left(dName, dName.LastIndexOf(".".ToCharArray))
            End If
        End If
        Return dName
    End Function

    Public Shared Function Path(Filename As String) As String
        If Filename = "" Then Return ""
        Filename = RemoveQuotationMarks(Filename)
        If Filename.LastIndexOf("\") = -1 Then Return ""
        Return Strings.Left(Filename, Filename.LastIndexOf("\"))
    End Function

    Public Shared Function Arguments(Filename As String) As String
        Dim Pos As Integer = Filename.IndexOf(Chr(34))
        If Not Pos = -1 Then Filename = Filename.Substring(Pos + 1, Filename.Length - Pos - 1)
        Pos = Filename.IndexOf(Chr(34))
        If Not Pos = -1 Then
            Arguments = Filename.Substring(Pos + 1, Filename.Length - Pos - 1)
            If Len(Arguments) > 0 Then Arguments = Arguments.Substring(1, Arguments.Length - 1)
            Return Arguments
        End If
        Pos = Filename.IndexOf("/")
        If Not Pos = -1 Then
            Return Filename.Substring(Pos, Filename.Length - Pos)
        End If
        Return ""
    End Function

    Public Shared Function RemoveQuotationMarks(Text As String) As String
        Dim Pos As Integer = Text.IndexOf(Chr(34))
        If Not Pos = -1 Then Text = Text.Substring(Pos + 1, Len(Text) - Pos - 1)
        Pos = Text.IndexOf(Chr(34))
        If Not Pos = -1 Then Text = Text.Substring(0, Pos)
        Pos = Text.IndexOf("//")
        If Pos = -1 Then
            Pos = Text.IndexOf("/")
            If Not Pos = -1 Then Text = Text.Substring(0, Pos - 1)
        End If
        Return Text
    End Function

    Public Shared Function Exist(cesta As String, Optional SystemFile As Boolean = False) As Boolean
        If cesta = "" Then Return False
        cesta = RemoveQuotationMarks(cesta)
        Try
            Dim exFile As New System.IO.FileInfo(cesta)
            If exFile.Exists = False Then Return False
            If SystemFile = False AndAlso exFile.Attributes = 6 Or exFile.Attributes = 38 Then Return False
        Catch
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Delete(ByVal Filename As String, ByVal Dustbin As Boolean,
                             Optional ByVal Inform As Boolean = True,
                             Optional ByVal DelReadOnly As Boolean = False) As Boolean
        If Filename = "" Then Return False
        Filename = RemoveQuotationMarks(Filename)
        Dim sTitle As String = "File remove "
        Try
            Dim delFile As New System.IO.FileInfo(Filename)
            If delFile.Exists = False Then Return True
            If delFile.IsReadOnly Then
                If DelReadOnly Then
                    delFile.Attributes = IO.FileAttributes.Normal
                Else
                    If Inform Then MessageBox.Show("File " & Name(Filename) & " is inaccessible.", sTitle)
                    Return False
                End If
            End If
            If Dustbin Then
                My.Computer.FileSystem.DeleteFile(Filename, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Else
                delFile.Delete()
            End If
        Catch
            If Inform Then MessageBox.Show("File " & Name(Filename) & " is inaccessible.", sTitle)
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Copy(source As String, destination As String, Optional overwrite As Boolean = True) As Boolean
        If Exist(source, True) Then
            If myFolder.Exist(Path(destination), True) Then
                If overwrite = False Then
                    Dim a As Integer
                    Dim newFile As String = destination
                    Do Until Not myFile.Exist(newFile)
                        a += 1
                        newFile = destination & "(" & CStr(a) & ")"
                    Loop
                    destination = newFile
                End If
                If Delete(destination, False) Then
                    Try
                        FileCopy(source, destination)
                        Return True
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If
        Return False
    End Function

    Public Shared Function Move(ByVal source As String, ByVal destination As String) As Boolean
        If Copy(source, destination) Then
            Delete(source, False)
            Return True
        End If
        Return False
    End Function

    Public Shared Sub Launch(Wnd As Window, ByVal Filename As String, Optional ByVal Admin As Boolean = False, Optional ErrMsg As String = "")
        Dim newProcess As New System.Diagnostics.ProcessStartInfo()
        newProcess.FileName = RemoveQuotationMarks(Filename)
        newProcess.Arguments = Arguments(Filename)
        newProcess.WorkingDirectory = Path(newProcess.FileName)
        If Admin Then newProcess.Verb = "runas"
        newProcess.CreateNoWindow = True
        Try
            Process.Start(newProcess)
        Catch Ex As Exception
            If Not Err.Number = 5 Then MessageBox.Show(Filename + NR + NR + If(ErrMsg = "", Ex.Message, ErrMsg), "Launch failed")
        End Try
    End Sub

End Class