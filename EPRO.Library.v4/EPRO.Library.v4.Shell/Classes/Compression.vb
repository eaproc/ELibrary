'Requires 4.0 FrameWork

'Imports System.IO.Compression
Imports System.IO
Imports System.IO.Packaging
Imports EPRO.Library.v3._5.EIO

'
'
'   Currently working but not the best compression in terms of size
'
'



''' <summary>
''' Contains Static Methods for compressing and extracting files.
''' </summary>
''' <remarks></remarks>
Public Class Compression
    'Make sure you have administrator's rights to run this code

#Region "Properties"
    ''' <summary>
    ''' I will use this to address space in compression and extraction
    ''' </summary>
    ''' <remarks></remarks>
    Const Replace_Space As String = "_(o_O)_"

#End Region

#Region "General Declarations"

#Region "Ignore"

    '' ''' <summary>
    '' ''' Gets file and zip it. It preserves the hierachy
    '' ''' </summary>
    '' ''' <param name="_Files"></param>
    '' ''' <param name="_FileName">Zip File Name</param>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Public Shared Function CreateZipFile(ByVal _FileName As String,
    ''                              ByVal ParamArray _Files() As String) As Boolean
    ''    Try

    ''        Dim _Zipper As New Ionic.Zip.ZipFile(_FileName)
    ''        With _Zipper

    ''            .AddFiles(_Files.AsEnumerable, False,
    ''                      FileIO.FileSystem.GetParentPath(_Files(0))
    ''                      )
    ''            .Save()
    ''        End With
    ''    Catch ex As Exception
    ''        Return False
    ''    End Try

    ''    Return True
    ''End Function


    '' ''' <summary>
    '' ''' Gets file and zip it.  [Only Files in Current Level]
    '' ''' </summary>
    '' ''' <param name="_Files">With their full Path</param>
    '' ''' <param name="_FileName">Zip File Name [Full Path]</param>
    '' ''' <returns></returns>
    '' ''' <remarks>Using Imports System.IO.Packaging</remarks>
    ''Public Shared Function CreateZipFile(ByVal _FileName As String,
    ''                              ByVal ParamArray _Files() As String) As Boolean
    ''    Try

    ''        'No Appending Allowed here so kill it if already exists
    ''        EIODeleteFileIfExists(_FileName)

    ''        'Open the zip file if it exists, else create a new one 
    ''        Dim zip As Package = ZipPackage.Open(_FileName, _
    ''             IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)

    ''        For Each _file As String In _Files
    ''            AddToArchive(zip, _file)

    ''        Next

    ''        zip.Close() 'Close the zip file


    ''        ' ''Dim _Zipper As New Ionic.Zip.ZipFile(_FileName)
    ''        ' ''With _Zipper

    ''        ' ''    .AddFiles(_Files.AsEnumerable, False,
    ''        ' ''              FileIO.FileSystem.GetParentPath(_Files(0))
    ''        ' ''              )
    ''        ' ''    .Save()
    ''        ' ''End With
    ''    Catch ex As Exception
    ''        Return False
    ''    End Try

    ''    Return True
    ''End Function


    '' '' ''' <summary>
    '' '' ''' Zip a whole directory with all it's contents [It doesnt Includes the Parent Directory]
    '' '' ''' </summary>
    '' '' ''' <param name="_FileName"></param>
    '' '' ''' <param name="DirectoryPath"></param>
    '' '' ''' <returns></returns>
    '' '' ''' <remarks>Using Iconic</remarks>
    '' ''Public Shared Function CreateZipFile(ByVal _FileName As String,
    '' ''                            ByVal DirectoryPath As String) As Boolean
    '' ''    Try

    '' ''        If File.Exists(_FileName) Then File.Delete(_FileName)

    '' ''        Dim _Zipper As New Ionic.Zip.ZipFile(_FileName)
    '' ''        With _Zipper
    '' ''            .AddDirectory(DirectoryPath, EIOGetDirectoryName(DirectoryPath))
    '' ''            .Save()
    '' ''        End With
    '' ''    Catch ex As Exception
    '' ''        Return False
    '' ''    End Try

    '' ''    Return True
    '' ''End Function


    ' '' ''' <summary>
    ' '' ''' Zip a whole directory with all it's contents Without Including the directory name itself
    ' '' ''' </summary>
    ' '' ''' <param name="_FileName"></param>
    ' '' ''' <param name="DirectoryPath"></param>
    ' '' ''' <returns></returns>
    ' '' ''' <remarks>Using Iconic</remarks>
    ' ''Public Shared Function CreateZipFile(ByVal _FileName As String,
    ' ''                            ByVal DirectoryPath As String,
    ' ''                            ByVal DoNotIncludeTheParentDirectory As Boolean) As Boolean


    ' ''    If Not DoNotIncludeTheParentDirectory Then Return CreateZipFile(_FileName, DirectoryPath)

    ' ''    Try

    ' ''        If File.Exists(_FileName) Then File.Delete(_FileName)

    ' ''        Dim _Zipper As New Ionic.Zip.ZipFile(_FileName)
    ' ''        With _Zipper



    ' ''            REM First add files
    ' ''            Dim sFile As List(Of String) = FileIO.FileSystem.GetFiles(DirectoryPath).ToList

    ' ''            For Each _File As String In sFile

    ' ''                .AddFile(_File, "")

    ' ''            Next


    ' ''            Dim sDir As List(Of String) = FileIO.FileSystem.GetDirectories(DirectoryPath).ToList

    ' ''            For Each _dir As String In sDir


    ' ''                .AddDirectory(_dir, EIOGetDirectoryName(_dir))


    ' ''            Next

    ' ''            .Save()
    ' ''        End With
    ' ''    Catch ex As Exception
    ' ''        Return False
    ' ''    End Try

    ' ''    Return True
    ' ''End Function



    '' ''' <summary>
    '' ''' Add a file to a zip package
    '' ''' </summary>
    '' ''' <param name="zip"></param>
    '' ''' <param name="fileToAdd">File Full Path</param>
    '' ''' <remarks>Using Imports System.IO.Packaging</remarks>
    ''Private Shared Sub AddToArchive(ByVal zip As Package, _
    ''                    ByVal fileToAdd As String)

    ''    'Replace spaces with an underscore (_) 
    ''    Dim uriFileName As String = fileToAdd.Replace(" ", "_")

    ''    'A Uri always starts with a forward slash "/" 
    ''    Dim zipUri As String = String.Concat("/", _
    ''               IO.Path.GetFileName(uriFileName))

    ''    Dim partUri As New Uri(zipUri, UriKind.Relative)
    ''    Dim contentType As String = _
    ''               Net.Mime.MediaTypeNames.Application.Zip

    ''    'The PackagePart contains the information: 
    ''    ' Where to extract the file when it's extracted (partUri) 
    ''    ' The type of content stream (MIME type):  (contentType) 
    ''    ' The type of compression:  (CompressionOption.Normal)   
    ''    Dim pkgPart As PackagePart = zip.CreatePart(partUri, _
    ''               contentType, CompressionOption.Normal)

    ''    'Read all of the bytes from the file to add to the zip file 
    ''    Dim bites As Byte() = File.ReadAllBytes(fileToAdd)

    ''    'Compress and write the bytes to the zip file 
    ''    pkgPart.GetStream().Write(bites, 0, bites.Length)

    ''End Sub

    '' ''' <summary>
    '' ''' Unzip File to the specified directories with a xml file [Content_Types] listing the contents of the zip
    '' ''' </summary>
    '' ''' <param name="_FileName"></param>
    '' ''' <param name="DirectoryPath"></param>
    '' ''' <returns></returns>
    '' ''' <remarks>Using Iconic</remarks>
    ''Public Shared Function UnZipFilesToDirectory(ByVal _FileName As String,
    ''                            ByVal DirectoryPath As String) As Boolean

    ''    Try

    ''        Dim _Zipper As ZipFile = ZipFile.Read(_FileName)

    ''        With _Zipper
    ''            ' For Each e As ZipEntry In _Zipper
    ''            ' e.Extract(DirectoryPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
    ''            .ExtractAll(DirectoryPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
    ''            'Next
    ''        End With
    ''    Catch ex As Exception
    ''        Return False
    ''    End Try

    ''    Return True
    ''End Function


#End Region


#Region "Using Imports System.IO.Packaging"

#Region "Private"

    ''' <summary>
    ''' Adds File to an open Zip Package
    ''' </summary>
    ''' <param name="zip">The zip Package</param>
    ''' <param name="FilePath">File to Add</param>
    ''' <param name="RelativeDirectoryPath">The relative Directory Path [Without the File Name]</param>
    ''' <remarks>Using Imports System.IO.Packaging</remarks>
    Private Shared Function AddFileToZipPackage(ByVal zip As ZipPackage, ByVal FilePath As String,
                                                     Optional ByVal RelativeDirectoryPath As String = vbNullString) As Boolean
        Try




            Dim partUri As New Uri(String.Format("{0}/{1}", RelativeDirectoryPath, getFileName(FilePath)).Replace(" ", Replace_Space), UriKind.Relative)
            Dim contentType As String = _
                       Net.Mime.MediaTypeNames.Application.Zip

            'The PackagePart contains the information: 
            ' Where to extract the file when it's extracted (partUri) 
            ' The type of content stream (MIME type):  (contentType) 
            ' The type of compression:  (CompressionOption.Normal)   
            Dim pkgPart As PackagePart = zip.CreatePart(partUri, _
                       contentType, CompressionOption.Normal)

            ''Dim pkgPart As PackagePart = zip.CreatePart(partUri, _
            ''          contentType)

            Dim bytes() As Byte = File.ReadAllBytes(FilePath)

            pkgPart.GetStream().Write(bytes, 0, bytes.Length)


        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Adds File to an open Zip Package
    ''' </summary>
    ''' <param name="zip">The zip Package</param>
    ''' <param name="RelativeDirectoryPath">The relative Directory Path [Without the File Name]</param>
    ''' <remarks>Using Imports System.IO.Packaging</remarks>
    Private Shared Function AddNullFileToZipPackage(ByVal zip As ZipPackage, ByVal RelativeDirectoryPath As String) As Boolean
        Try


            Const NullFileName As String = "Null_File_____Thumb.xml"

            REM I added the replace because the relative directory could contain space
            Dim partUri As New Uri(String.Format("{0}/{1}", RelativeDirectoryPath, NullFileName).Replace(" ", Replace_Space), UriKind.Relative)
            Dim contentType As String = _
                       Net.Mime.MediaTypeNames.Application.Zip

            'The PackagePart contains the information: 
            ' Where to extract the file when it's extracted (partUri) 
            ' The type of content stream (MIME type):  (contentType) 
            ' The type of compression:  (CompressionOption.Normal)   
            zip.CreatePart(partUri, _
                       contentType, CompressionOption.Normal)


        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function


    ''' <summary>
    ''' Adds the directory to Package
    ''' </summary>
    ''' <param name="zip">The Zip Package</param>
    ''' <param name="DirectoryPath">The Directory to Add Full Path</param>
    ''' <param name="RelativePath">Accumulated Relative Path Along the Tree</param>
    ''' <remarks></remarks>
    Private Shared Function AddDirectoryToZipPackage(ByVal zip As ZipPackage, ByVal DirectoryPath As String,
                                                     Optional ByVal RelativePath As String = vbNullString) As Boolean

        ''REM With this method empty folders will not be added
        ''Const CurrentDirectoryConfig As String = "C:\Windows\Temp\__Default__Names___.xml"
        ''Const SectionName As String = "NAMES"

        ''REM Delete Old Config
        ''If File.Exists(CurrentDirectoryConfig) Then File.Delete(CurrentDirectoryConfig)



        Try



            REM First add files
            Dim sFile As List(Of String) = FileIO.FileSystem.GetFiles(DirectoryPath).ToList

            REM Preserve the Directory .. Even thou it is empty
            If sFile.Count = 0 Then AddNullFileToZipPackage(zip, RelativePath)

            For Each _File As String In sFile

                If Not AddFileToZipPackage(zip, _File, RelativePath) Then Exit For

            Next


            Dim sDir As List(Of String) = FileIO.FileSystem.GetDirectories(DirectoryPath).ToList

            For Each _dir As String In sDir


                If Not AddDirectoryToZipPackage(zip, _dir, String.Format("{0}/{1}", RelativePath, GetDirectoryName(_dir))) Then Exit For


            Next


        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

#End Region


    ''' <summary>
    ''' Zip a whole directory with all it's contents [It doesnt Includes the Parent Directory]
    ''' </summary>
    ''' <param name="_FileName">The .zip File Full Path</param>
    ''' <param name="DirectoryPath">The Directory That contains the files and sub folders to zip</param>
    ''' <returns></returns>
    ''' <remarks>Using Imports System.IO.Packaging</remarks>
    Public Shared Function CreateZipFile(ByVal _FileName As String,
                                ByVal DirectoryPath As String,
                                Optional ByVal IncludeCurrentDirectoryName As Boolean = False) As Boolean

        REM This is creating without adding the initial Directory
        REM This Class is strict with name conventions .. no space
        REM for each folder there will a File Called __Default__Names___.xml
        REM it will contain the folder real name and all files real name

        Try

            If File.Exists(_FileName) Then File.Delete(_FileName)


            Using zip As Package = ZipPackage.Open(_FileName, _
              IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)

                If IncludeCurrentDirectoryName Then

                    Return AddDirectoryToZipPackage(CType(zip, ZipPackage), DirectoryPath, "/" & GetDirectoryName(DirectoryPath))

                Else
                    Return AddDirectoryToZipPackage(CType(zip, ZipPackage), DirectoryPath)

                End If

            End Using


        Catch ex As Exception
            Return False
        End Try


    End Function


    ''' <summary>
    ''' Gets file and zip it.  [Only Files in Current Level]
    ''' </summary>
    ''' <param name="_Files">With their full Path</param>
    ''' <param name="_FileName">Zip File Name [Full Path]</param>
    ''' <returns></returns>
    ''' <remarks>Using Imports System.IO.Packaging</remarks>
    Public Shared Function CreateZipFile(ByVal _FileName As String,
                                  ByVal ParamArray _Files() As String) As Boolean
        Try

            If File.Exists(_FileName) Then File.Delete(_FileName)


            Using zip As Package = ZipPackage.Open(_FileName, _
              IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)


                For Each _File As String In _Files


                    AddFileToZipPackage(CType(zip, ZipPackage), _File)

                Next

            End Using


        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function


    ''' <summary>
    ''' Extract the Zip File to the Specified Folder [Note: Then Zip File Must have been created by System.IO.Packaging] 
    ''' </summary>
    ''' <param name="ExtractToFolderPath">The Diretory to which the files and folders will be extracted to</param>
    ''' <param name="ZipFileFullPath">The Zip File</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractZipFolder(ByVal ZipFileFullPath As String, ByVal ExtractToFolderPath As String) As Boolean
        REM This extraction only works if the zip file was created with the same method [Packaging]
        REM Because it needs the xml file inside to read the package



        Try

            If Not Directory.Exists(ExtractToFolderPath) Then Directory.CreateDirectory(ExtractToFolderPath)

            ' Open the Package.
            ' ('using' statement insures that 'package' is
            '  closed and disposed when it goes out of scope.)
            Using package As Package = package.Open(ZipFileFullPath, FileMode.Open, FileAccess.Read)
                For Each part As PackagePart In package.GetParts

                    Dim targetPath As String = ExtractToFolderPath & part.Uri.ToString().Replace("/", System.IO.Path.DirectorySeparatorChar).Replace(Replace_Space, " ")

                    If Not Directory.Exists(FileIO.FileSystem.GetParentPath(targetPath)) Then
                        Directory.CreateDirectory(FileIO.FileSystem.GetParentPath(targetPath))
                    End If

                    Dim bytes(CInt(part.GetStream().Length)) As Byte


                    part.GetStream().Read(bytes, 0, bytes.Length)

                    File.WriteAllBytes(targetPath, bytes)

                Next

            End Using ' end:using(Package package) - Close & dispose package.


        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

#End Region


#End Region



End Class
