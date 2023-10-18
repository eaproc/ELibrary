Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary


Imports ELibrary.Standard.VB.Modules

''' <summary>
''' Static Class for I/O functions
''' </summary>
''' <remarks></remarks>
Public Class EIO





    Public Const WINDOWS_TEMP_FOLDER As String = "C:\Windows\Temp"






    ''' <summary>
    ''' Try to empty a directory withought deleting the directory itself
    ''' </summary>
    ''' <param name="direc"></param>
    ''' <remarks></remarks>
    Public Shared Sub DeleteAllFromDirectory(ByVal direc As String)


        Dim SubDirs() As String = Directory.GetDirectories(direc)
        Dim SubFiles As String() = Directory.GetFiles(direc)


        If SubDirs IsNot Nothing Then

            For Each subDir As String In SubDirs
                Try

                    Directory.Delete(subDir, True)


                Catch ex As Exception

                End Try

            Next


        End If

        If SubFiles IsNot Nothing Then

            For Each subFile As String In SubFiles
                Try

                    File.Delete(subFile)


                Catch ex As Exception

                End Try

            Next


        End If


    End Sub


    '''' <summary>
    '''' Deletes a directory and all its contents if exist. It ignores error on delete
    '''' </summary>
    '''' <param name="DirectoryPath"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function DeleteLongPathDirectory(ByVal DirectoryPath As String) As Boolean


    '    Try

    '        With New MicrosoftOS.CommandPrompt(False, False)

    '            Return .Execute(string.Format("rmdir /S /Q ""{0}"" ", DirectoryPath))

    '        End With


    '    Catch ex As Exception
    '        Return False
    '    End Try

    '    Return True
    'End Function

    ''' <summary>
    ''' Deletes a directory and all its contents if exist. It ignores error on delete
    ''' </summary>
    ''' <param name="DirectoryPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteDirectory(ByVal DirectoryPath As String) As Boolean


        Try
            If Directory.Exists(
                                   DirectoryPath
                                 ) Then
                Directory.Delete(DirectoryPath, True)
            End If

        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Deletes a directory and all its contents if exist and recreate it as empty directory. Creates new if it doesnt exist
    ''' </summary>
    ''' <param name="DirectoryPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteAndRecreateDirectory(ByVal DirectoryPath As String) As Boolean


        Try
            If Directory.Exists(
                                   DirectoryPath
                                 ) Then
                Directory.Delete(
                                           DirectoryPath,
                True
                                            )
            End If


            Directory.CreateDirectory(
                               DirectoryPath
                                )
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Deletes a file if it exists
    ''' </summary>
    ''' <param name="_FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteFileIfExists(ByVal _FileName As String) As Boolean
        Try

            If File.Exists(_FileName) Then File.Delete(_FileName)


        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function






#Region "Path Naming"




    ''' <summary>
    ''' Gets the extension of a file. It returns (.) with the extension name
    ''' </summary>
    ''' <param name="__FileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileExtension(ByVal __FileName As String) As String

        Try


            Return IO.Path.GetExtension(__FileName)


        Catch ex As Exception

        End Try

        Return vbNullString
    End Function



    ''' <summary>
    ''' Fetches the directory name.. using backward slash. Only the current directoy like [MyFolder] is returned not full directory
    ''' </summary>
    ''' <param name="DirectoryPath"></param>
    ''' <param name="IncludeSlash"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDirectoryName(ByVal DirectoryPath As String,
                                            Optional ByVal IncludeSlash As Boolean = False) As String


        Dim Dir_Path() As String = DirectoryPath.Split("\")


        If IncludeSlash Then

            Return "\" & Dir_Path(Dir_Path.Length - 1)

        Else

            Return Dir_Path(Dir_Path.Length - 1)


        End If


    End Function

    ''' <summary>
    ''' Get's a directory full path from a file full path
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDirectoryFullPath(ByVal FilePath As String) As String
        Try
            Return Path.GetDirectoryName(FilePath)

        Catch ex As Exception
            Return vbNullString
        End Try

    End Function

    ''' <summary>
    ''' Get a file name without extension
    ''' </summary>
    ''' <param name="__FileName">a File Name without the full path</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileNameWithoutExtension(ByVal __FileName As String) As String
        Dim fName As String = IO.Path.GetFileName(__FileName)

        Return Path.GetFileNameWithoutExtension(__FileName)
    End Function

    ''' <summary>
    ''' Get the file name in the provided uri
    ''' </summary>
    ''' <param name="uri"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileName(ByVal uri As Uri) As String

        Return Path.GetFileName(uri.AbsolutePath)

        Return vbNullString
    End Function


    ''' <summary>
    ''' Get the file name in the provided uri
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileName(ByVal URL As String) As String

        Return Path.GetFileName(URL)


        Return vbNullString
    End Function






    '   Check the Modeling file for more explanation


    ''' <summary>
    ''' get a suggested file full path name. like c:\...File.txt, c:\...File___1.txt
    ''' </summary>
    ''' <param name="pIntendedFileFullPath">the intended file name</param>
    ''' <param name="pIncrementSeparator">if you want a separator between filename and increment. NB: Increment must be file naming compatible.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SuggestFileUniqueFullFilePathName(ByVal pIntendedFileFullPath As String,
                                                             Optional ByVal pIncrementSeparator As String = "___") As String
        Return SuggestFileUniqueFullFilePathName(EIO.GetDirectoryFullPath(pIntendedFileFullPath),
                                                 EIO.GetFileExtension(pIntendedFileFullPath),
                                                 pIncrementSeparator
                                                 )
    End Function


    ''' <summary>
    ''' get a suggested file full path name. like c:\...File.txt, c:\...File___1.txt
    ''' </summary>
    ''' <param name="pIntendedFileFolderWithBackSlash">The directory you wish to create the file in with back slash C:\</param>
    ''' <param name="intendedFileName">the intended file name</param>
    ''' <param name="IncrementSeparator">if you want a separator between filename and increment. NB: Increment must be file naming compatible.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SuggestFileUniqueFullFilePathName(ByVal pIntendedFileFolderWithBackSlash As String,
                                                             ByVal intendedFileName As String,
                                                              ByVal IncrementSeparator As String) As String


        REM Path.getExtension returns dot as well like .csv

        If Not Directory.Exists(pIntendedFileFolderWithBackSlash) Then Return string.Empty
        If Not pIntendedFileFolderWithBackSlash.EndsWith("\") Then pIntendedFileFolderWithBackSlash &= "\"

        If Not File.Exists(pIntendedFileFolderWithBackSlash & intendedFileName) Then _
                Return pIntendedFileFolderWithBackSlash & intendedFileName


        Dim Increments() As String =
                Directory.GetFiles(pIntendedFileFolderWithBackSlash,
                                 string.Format("{0}{1}*{2}", Path.GetFileNameWithoutExtension(intendedFileName),
                                               IncrementSeparator, Path.GetExtension(intendedFileName)
                                               )
                               )

        If Increments Is Nothing OrElse Increments.Length = 0 Then Return string.Format("{0}{1}{2}1{3}",
                                            pIntendedFileFolderWithBackSlash,
                                            Path.GetFileNameWithoutExtension(intendedFileName),
                                               IncrementSeparator, Path.GetExtension(intendedFileName)
                                               )


        REM It is always sorted according to name ascending order by default [NOT GAURANTEED]
        '   Array.Sort(Increments)  REM Gaurantee sorting

        Increments = Increments.ToList().OrderByDescending(Function(x) x).ToArray()
        Dim pTopmostIncrementFileFullPath As String = Increments.First()


        REM Just the filename without folder path and extension
        Dim pIntendedFileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(intendedFileName)



        '   Fetching the incrementValue
        Dim pIncrement As Int32 = 0

        '   
        If pIntendedFileNameWithoutExtension.equalsIgnoreCase(Path.GetFileNameWithoutExtension(pTopmostIncrementFileFullPath)) Then
            pIncrement = 1

        Else

            Dim pTopmostIncrementFileName = Path.GetFileNameWithoutExtension(pTopmostIncrementFileFullPath)
            Dim pDifference = pTopmostIncrementFileName.Substring(pIntendedFileNameWithoutExtension.Length,
                                                                       pTopmostIncrementFileName.Length - pIntendedFileNameWithoutExtension.Length
                                                                       )
            pIncrement = InputsParsing.TextParsing.parseOutIntegers(pDifference).ToInt32() + 1

        End If


        Return string.Format("{0}{1}{2}{3}{4}",
                                             pIntendedFileFolderWithBackSlash,
                                            Path.GetFileNameWithoutExtension(intendedFileName),
                                               IncrementSeparator, pIncrement,
                                               Path.GetExtension(intendedFileName)
                                               )


        Return string.Empty

    End Function


    '''' <summary>
    '''' Scan a directory for available file name you can use and returns full path of the available file name
    '''' </summary>
    '''' <param name="FolderPath"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function getAvailableFileName(ByVal FolderPath As String,
    '                                            Optional ByVal suggestedFileName As String = "File.dat") As String

    '    Dim iCount As Long = 0


    '    If Not Directory.Exists(FolderPath) Then Return ""

    '    Dim __ext As String = New FileInfo(suggestedFileName).Extension
    '    Dim sName As String = getFileNameWithoutExtension(suggestedFileName)
    '    Dim chkFile As String = sName


    '    REM Last slash if it exist
    '    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)


    '    While True

    '        If File.Exists(
    '            string.Format(
    '                "{0}\{1}{2}", FolderPath, chkFile, __ext
    '                )
    '            ) Then

    '            iCount += 1

    '            chkFile = string.Format(
    '                                "{0}_{1}", sName, iCount
    '                                )
    '        Else

    '            Return string.Format(
    '                "{0}\{1}{2}", FolderPath, chkFile, __ext
    '                )

    '        End If

    '    End While

    '    Return vbNullString

    'End Function




#End Region






    ''' <summary>
    ''' Return the File path with the file name for the type that was sent
    ''' </summary>
    ''' <param name="objType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Get_Code_File_StartUp_File_FullPathName(objType As System.Type) As String
        Return objType.Assembly.Location
    End Function

    ''' <summary>
    ''' Return the Only File path WITHOUT the file name for the type that was sent
    ''' </summary>
    ''' <param name="objType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Get_Code_File_StartUp_Path(ByVal objType As System.Type) As String

        Return System.IO.Directory.GetParent(objType.Assembly.Location).ToString

    End Function

    ''' <summary>
    ''' Loads a File (.exe,.dll) and reads it Assembly for GUID attributes
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Get_GUID_From_File(ByVal FilePath As String) As String
        Try

            Dim fLoadedByte As Byte() = System.IO.File.ReadAllBytes(FilePath)

            Dim asy As System.Reflection.Assembly = System.Reflection.Assembly.Load(fLoadedByte)

            Return (
              CType(
                  asy.GetCustomAttributes(
                                   GetType(System.Runtime.InteropServices.GuidAttribute), True
                                                                                      )(0), System.Runtime.InteropServices.GuidAttribute
                                                                                  ).Value
                                                                                  )
        Catch ex As Exception
            Return string.Empty
        End Try
    End Function



    ''' <summary>
    ''' Confirms if a directory is writable by this current application. [If it has the access level]
    ''' </summary>
    ''' <param name="FolderPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsFolderWritableAndReadable(ByVal FolderPath As String) As Boolean

        Try

            File.WriteAllText(
                    string.Format("{0}\{1}", FolderPath, "Test.dll"), "Test"
                    )

            If File.ReadAllText(
                 string.Format("{0}\{1}", FolderPath, "Test.dll")
                 ) = "Test" Then

                DeleteFileIfExists(
                     string.Format("{0}\{1}", FolderPath, "Test.dll")
                     )
                Return True
            End If
        Catch ex As Exception

        End Try

        Return False

    End Function


    '''' <summary>
    '''' Checks if this a file by checking File.GetAttributes(__FileName) NOT Equals FileAttributes.Directory
    '''' </summary>
    '''' <param name="__FileName">a File Name without the full path</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function isThis_a_file(ByVal __FileName As String) As Boolean

    '    Try
    '        Dim p = File.GetAttributes(__FileName)
    '        Return p <> FileAttributes.Directory AndAlso
    '            p <> (FileAttribute.Directory Or FileAttribute.Archive)
    '    Catch ex As Exception
    '        MyLogFile.Print(ex)
    '        Return False
    '    End Try


    'End Function






#Region "Serialization"


#Region "SOAP"

    ''   SOAP Serialization is great for deep serialization.
    ''   It requires implementation of ISerializable and/or Serializable attribute
    ''   
    ''   It is same as Binary Serializable with the addition of portability
    ''   It is portable because it only depends on the class.


    '''' <summary>
    '''' Writes Serializable Object to File. Throws all exceptions
    '''' </summary>
    '''' <param name="obj"></param>
    '''' <param name="FileFullPath"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function WriteToFileSOAP(ByVal obj As Object, ByVal FileFullPath As String) As Boolean
    '    REM Create path first
    '    If Not Directory.Exists(Path.GetDirectoryName(FileFullPath)) Then _
    '        Directory.CreateDirectory(Path.GetDirectoryName(FileFullPath))

    '    Dim fs As New FileStream(FileFullPath, FileMode.Create)
    '    Dim objSerializer As New SoapFormatter()
    '    objSerializer.Serialize(fs, obj)

    '    fs.Close()

    '    Return True

    'End Function

    '''' <summary>
    '''' Reads a serialized object from file. Throws all exceptions
    '''' </summary>
    '''' <typeparam name="T"></typeparam>
    '''' <param name="FileFullPath"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function ReadFromFileSOAP(Of T)(ByVal FileFullPath As String) As T
    '    Dim fs As New FileStream(FileFullPath, FileMode.Open, FileAccess.Read)
    '    Dim objSerializer As New SoapFormatter()

    '    Dim rst As T = CType(objSerializer.Deserialize(fs), T)

    '    fs.Close()

    '    Return rst

    'End Function

#End Region




#Region "Binary"

    '   Binary Serialization is great for deep serialization
    '   It requires implementation of ISerializable and/or Serializable attribute
    '   
    '   It is NOT portable because it is deep and depends fully on the Namespace.Class that wrote it.


    ''' <summary>
    ''' Writes Serializable Object to File. Throws all exceptions
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="FileFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteToFile(ByVal obj As Object, ByVal FileFullPath As String) As Boolean
        REM Create path first
        If Not Directory.Exists(Path.GetDirectoryName(FileFullPath)) Then _
            Directory.CreateDirectory(Path.GetDirectoryName(FileFullPath))

        Dim fs As New FileStream(FileFullPath, FileMode.Create)
        Dim objSerializer As New BinaryFormatter()
        objSerializer.Serialize(fs, obj)

        fs.Close()

        Return True

    End Function

    ''' <summary>
    ''' Reads a serialized object from file. Throws all exceptions
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="FileFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFromFile(Of T)(ByVal FileFullPath As String) As T
        Dim fs As New FileStream(FileFullPath, FileMode.Open, FileAccess.Read)
        Dim objSerializer As New BinaryFormatter()

        Dim rst As T = CType(objSerializer.Deserialize(fs), T)

        fs.Close()

        Return rst

    End Function

#End Region


#Region "XML"


    '   XML Serialization is good for shallow serialization.
    '   It does NOT require Serializable attribute or implementation
    '   It ONLY requires a default constructor (Parameterless)
    '
    '   It serialize only property declared public read and write
    '
    '   It is portable because it only depends on the class that calls it



    ''' <summary>
    ''' Writes Serializable Object to File. Throws all exceptions
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="FileFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteToFileXML(Of T)(ByVal obj As T, ByVal FileFullPath As String) As Boolean
        REM Create path first
        If Not Directory.Exists(Path.GetDirectoryName(FileFullPath)) Then _
            Directory.CreateDirectory(Path.GetDirectoryName(FileFullPath))

        Dim fs As New FileStream(FileFullPath, FileMode.Create)
        Dim objSerializer As New Xml.Serialization.XmlSerializer(obj.GetType())
        objSerializer.Serialize(fs, obj)

        fs.Close()

        Return True

    End Function

    ''' <summary>
    ''' Reads a serialized object from file. Throws all exceptions
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="FileFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFromFileXML(Of T)(ByVal FileFullPath As String) As T
        Dim fs As New FileStream(FileFullPath, FileMode.Open, FileAccess.Read)
        Dim objSerializer As New Xml.Serialization.XmlSerializer(GetType(T))
        Dim reader As Xml.XmlReader = Xml.XmlReader.Create(fs)

        Dim rst As T = CType(objSerializer.Deserialize(reader), T)

        fs.Close()

        Return rst

    End Function




#End Region




#End Region





    ''' <summary>
    ''' Safely copies a file from one destination to another by making sure destination folder exists. Throws other exception
    ''' </summary>
    ''' <param name="src"></param>
    ''' <param name="dst"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CopyFile(ByVal src As String, ByVal dst As String,
                                         Optional overwrite As Boolean = True) As Boolean
        If Not Directory.Exists(GetDirectoryFullPath(dst)) Then Directory.CreateDirectory(GetDirectoryFullPath(dst))
        File.Copy(src, dst, overwrite)
        Return True
    End Function

    ''' <summary>
    ''' Safely copies a file from one destination to another by making sure destination folder exists. Throws other exception
    ''' </summary>
    ''' <param name="src"></param>
    ''' <param name="dstDir"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CopyFileToDir(ByVal src As String, ByVal dstDir As String,
                                         Optional overwrite As Boolean = True) As Boolean
        If Not Directory.Exists(dstDir) Then Directory.CreateDirectory(dstDir)
        Return CopyFile(src, string.Format("{0}\{1}", dstDir, GetFileName(src)), overwrite)
    End Function

End Class
