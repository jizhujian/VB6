''' <summary>
''' 文件操作函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(File.ClassId, File.InterfaceId, File.EventsId)> _
Public Class File

    ''' <summary>
    ''' COM注册必须
    ''' </summary>
    ''' <remarks></remarks>
    Public Const ClassId As String = "989e3503-32d9-4feb-8da2-0b25e6981578"
    Public Const InterfaceId As String = "ada71512-51eb-4e53-a23d-10517d9b2450"
    Public Const EventsId As String = "61becf92-7f30-49a9-91fb-ac7703844051"

    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' 字符编码
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TextEncoding
        ''' <summary>
        ''' 操作系统的当前 ANSI 代码页的编码
        ''' </summary>
        ''' <remarks></remarks>
        [Default] = 0
        ''' <summary>
        ''' Little-Endian 字节顺序的 UTF-16 格式的编码
        ''' </summary>
        ''' <remarks></remarks>
        Unicode
        ''' <summary>
        ''' ASCII（7 位）字符集的编码
        ''' </summary>
        ''' <remarks></remarks>
        ASCII
        ''' <summary>
        ''' UTF-8 格式的编码
        ''' </summary>
        ''' <remarks></remarks>
        UTF8
        ''' <summary>
        ''' UTF-7 格式的编码
        ''' </summary>
        ''' <remarks></remarks>
        UTF7
        ''' <summary>
        ''' Little-Endian 字节顺序的 UTF-32 格式的编码
        ''' </summary>
        ''' <remarks></remarks>
        UTF32
        ''' <summary>
        ''' Big-Endian 字节顺序的 UTF-16 格式的编码
        ''' </summary>
        ''' <remarks></remarks>
        BigEndianUnicode
    End Enum

    ''' <summary>
    ''' 将COM字符编码类型转换NET字符编码类型
    ''' </summary>
    ''' <param name="encoding"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function GetTextEncoding(ByVal encoding As TextEncoding) As System.Text.Encoding
        Select Case encoding
            Case TextEncoding.Unicode
                Return System.Text.Encoding.Unicode
            Case TextEncoding.ASCII
                Return System.Text.Encoding.ASCII
            Case TextEncoding.UTF8
                Return System.Text.Encoding.UTF8
            Case TextEncoding.UTF7
                Return System.Text.Encoding.UTF7
            Case TextEncoding.UTF32
                Return System.Text.Encoding.UTF32
            Case TextEncoding.BigEndianUnicode
                Return System.Text.Encoding.BigEndianUnicode
            Case Else
                Return System.Text.Encoding.Default
        End Select
    End Function

    ''' <summary>
    ''' 返回当前系统的临时文件夹的路径
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTempPath() As String
        Return System.IO.Path.GetTempPath
    End Function

    ''' <summary>
    ''' 创建磁盘上唯一命名的零字节的临时文件并返回该文件的完整路径
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTempFileName() As String
        Return System.IO.Path.GetTempFileName
    End Function

    ''' <summary>
    ''' 打开一个文件，将文件的内容读入一个字节数组，然后关闭该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadAllBytes(ByVal path As String) As Byte()
        Return System.IO.File.ReadAllBytes(path)
    End Function

    ''' <summary>
    ''' 打开一个文件，使用指定的编码，读取文件的所有行到字符串数组，然后关闭该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="encoding"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadAllLines(ByVal path As String, Optional ByVal encoding As TextEncoding = TextEncoding.Default) As String()
        Return System.IO.File.ReadAllLines(path, GetTextEncoding(encoding))
    End Function

    ''' <summary>
    ''' 打开一个文件，使用指定的编码，读取文件的所有行到字符串，然后关闭该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="encoding"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadAllText(ByVal path As String, Optional ByVal encoding As TextEncoding = TextEncoding.Default) As String
        Return System.IO.File.ReadAllText(path, GetTextEncoding(encoding))
    End Function

    ''' <summary>
    ''' 创建一个新文件，在其中写入指定的字节数组，然后关闭该文件。如果目标文件已存在，则覆盖该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="bytes"></param>
    ''' <remarks></remarks>
    Public Sub WriteAllBytes(ByVal path As String, ByRef bytes() As Byte)
        System.IO.File.WriteAllBytes(path, bytes)
    End Sub

    ''' <summary>
    ''' 创建一个新文件，使用指定的编码在其中写入指定的字符串数组，然后关闭该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="contents"></param>
    ''' <param name="encoding"></param>
    ''' <remarks></remarks>
    Public Sub WriteAllLines(ByVal path As String, ByRef contents As String(), Optional ByVal encoding As TextEncoding = TextEncoding.Default)
        System.IO.File.WriteAllLines(path, contents, GetTextEncoding(encoding))
    End Sub

    ''' <summary>
    ''' 创建一个新文件，使用指定的编码在其中写入指定的字符串，然后关闭文件。如果目标文件已存在，则覆盖该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="contents"></param>
    ''' <param name="encoding"></param>
    ''' <remarks></remarks>
    Public Sub WriteAllText(ByVal path As String, ByVal contents As String, Optional ByVal encoding As TextEncoding = TextEncoding.Default)
        System.IO.File.WriteAllText(path, contents, GetTextEncoding(encoding))
    End Sub

    ''' <summary>
    ''' 使用指定的编码将指定的字符串追加到文件中，如果文件还不存在则创建该文件
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="contents"></param>
    ''' <param name="encoding"></param>
    ''' <remarks></remarks>
    Public Sub AppendAllText(ByVal path As String, ByVal contents As String, Optional ByVal encoding As TextEncoding = TextEncoding.Default)
        System.IO.File.AppendAllText(path, contents, GetTextEncoding(encoding))
    End Sub

    ''' <summary>
    ''' 使用指定的编码向一个文件中追加文本行，然后关闭该文件。
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="contents"></param>
    ''' <param name="encoding"></param>
    ''' <remarks></remarks>
    Public Sub AppendAllLines(ByVal path As String, ByRef contents() As String, Optional ByVal encoding As TextEncoding = TextEncoding.Default)
        System.IO.File.AppendAllLines(path, contents, GetTextEncoding(encoding))
    End Sub

    ''' <summary>
    ''' 确定指定的文件是否存在
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExistsFile(ByVal path As String) As Boolean
        Return System.IO.File.Exists(path)
    End Function

    ''' <summary>
    ''' 删除指定的文件。如果指定的文件不存在，则不引发异常
    ''' </summary>
    ''' <param name="path"></param>
    ''' <remarks></remarks>
    Public Sub DeleteFile(ByVal path As String)
        System.IO.File.Delete(path)
    End Sub

    ''' <summary>
    ''' 返回指定路径字符串的目录信息
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDirectoryName(ByVal path As String) As String
        Return System.IO.Path.GetDirectoryName(path)
    End Function

    ''' <summary>
    ''' 返回指定路径字符串的文件名和扩展名
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileName(ByVal path As String) As String
        Return System.IO.Path.GetFileName(path)
    End Function

    ''' <summary>
    ''' 去除文件名中无效字符
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValidatedFileName(ByVal path As String) As String
        Return path.Replace("\", "").Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("""", "").Replace("<", "").Replace(">", "").Replace("|", "").Trim
    End Function

    ''' <summary>
    ''' 返回不具有扩展名的指定路径字符串的文件名
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileNameWithoutExtension(ByVal path As String) As String
        Return System.IO.Path.GetFileNameWithoutExtension(path)
    End Function

    ''' <summary>
    ''' 确定路径是否包括文件扩展名
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HasExtension(ByVal path As String) As Boolean
        Return System.IO.Path.HasExtension(path)
    End Function

    ''' <summary>
    ''' 返回指定的路径字符串的扩展名
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExtension(ByVal path As String) As String
        Return System.IO.Path.GetExtension(path)
    End Function

    ''' <summary>
    ''' 更改路径字符串的扩展名
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="extension">新的扩展名（有或没有前导句点）。指定 null 以从 path 移除现有扩展名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChangeExtension(ByVal path As String, ByVal extension As String) As String
        Return System.IO.Path.ChangeExtension(path, extension)
    End Function

    Public Function GetDesktopPath() As String
        Return System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
    End Function

    Public Function GetMyDocumentsPath() As String
        Return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
    End Function

    Public Sub CopyFile(ByVal sourceFileName As String, ByVal destFileName As String, ByVal overwrite As Boolean)
        System.IO.File.Copy(sourceFileName, destFileName, overwrite)
    End Sub

    Public Sub CreateDirectory(ByVal path As String)
        System.IO.Directory.CreateDirectory(path)
    End Sub

    Public Function ExistsDirectory(ByVal path As String) As Boolean
        Return System.IO.Directory.Exists(path)
    End Function

    Public Sub DeleteDirectory(ByVal path As String)
        System.IO.Directory.Delete(path, True)
    End Sub

    Public Function GetDirectories(ByVal path As String) As String()
        Return System.IO.Directory.GetDirectories(path)
    End Function

    Public Function GetFiles(ByVal path As String) As String()
        Return System.IO.Directory.GetFiles(path)
    End Function

    Public Function GetFilesByPattern(ByVal path As String, ByVal searchPattern As String) As String()
        Return System.IO.Directory.GetFiles(path, searchPattern)
    End Function

    Public Function Combine(ByVal path1 As String, ByVal path2 As String) As String
        Return System.IO.Path.Combine(path1, path2)
    End Function

End Class
