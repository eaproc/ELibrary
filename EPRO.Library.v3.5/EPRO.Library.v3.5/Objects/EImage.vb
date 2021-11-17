Imports System.Drawing
Imports EPRO.Library.v3._5.EIO

Namespace Objects

    Public Class EImage

        Public Shared Function valueOf(ByVal BytesArray() As Byte) As Image

            Try

                If BytesArray Is Nothing Then Return Nothing


                Using ms As New IO.MemoryStream(BytesArray, 0, BytesArray.Length)

                    ms.Write(BytesArray, 0, BytesArray.Length)

                    '   New Bitmap is to avoid "A generic error occurred in GDI+" when saving back
                    Return New Bitmap(Image.FromStream(ms, True))
                End Using




            Catch ex As Exception

                Return Nothing

            End Try

        End Function

        Public Shared Function valueOf(ByVal SysDBNULL As System.DBNull) As Image
            Return Nothing
        End Function

        Public Shared Function valueOf(ByVal pVal As Object) As Image
            If pVal Is Nothing Then Return Nothing
            If TypeOf pVal Is Byte() Then Return valueOf(CType(pVal, Byte()))
            If TypeOf pVal Is System.DBNull Then Return valueOf(CType(pVal, System.DBNull))

            Return Nothing REM We don't understand the format

        End Function







        ''' <summary>
        ''' To solve the issue with image saving not being able to decern it's format
        ''' </summary>
        ''' <param name="Img"></param>
        ''' <param name="FilePath">The real output is Bitmaps(bmp)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SaveImage_To_File(ByVal Img As Image, ByVal FilePath As String) As Boolean
            Try

                Dim newBmp As Bitmap = New Bitmap(Img)

                newBmp.Save(FilePath, Imaging.ImageFormat.Bmp)

                Return True
            Catch ex As Exception

            End Try

            Return False

        End Function










        ''' <summary>
        ''' Get Icon from application type
        ''' </summary>
        ''' <param name="objType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function get_Icon_From_Object_Type(ByVal objType As System.Type) As Icon

            Try
                Return Icon.ExtractAssociatedIcon(get_Code_File_StartUp_File_FullPathName(objType.GetType()))
            Catch ex As Exception
                Return Nothing
            End Try

        End Function









        Public Shared Function toByteArray(ByVal pImage As Image, pImageFormat As Imaging.ImageFormat) As Byte()
            Try
                Dim data As Byte() = Nothing
                Dim ms As New IO.MemoryStream()

                Dim pMemoryFormatWorked As Boolean = False
                Try
                    pImage.Save(ms, pImageFormat)
                    pMemoryFormatWorked = True
                Catch ex As Exception
                    Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image, pImageFormat As Imaging.ImageFormat)", ex)

                End Try

                If Not pMemoryFormatWorked Then
                    '   Try Bitmap
                    pImage.Save(ms, Imaging.ImageFormat.Bmp)

                End If

                data = ms.GetBuffer()
                ms = Nothing
                Return data
            Catch ex As Exception
                Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image, pImageFormat As Imaging.ImageFormat)", ex)

                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' It doesnt throw exception
        ''' </summary>
        ''' <param name="pImage"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toByteArray(ByVal pImage As Image) As Byte()
            Try
                If pImage Is Nothing Then Return Nothing
                Return toByteArray(pImage, pImage.RawFormat)

            Catch ex As Exception
                Modules.basMain.MyLogFile.Print("toByteArray(ByVal pImage As Image)", ex)
                Modules.basMain.MyLogFile.Print(ex)

                Return Nothing
            End Try
        End Function







#Region "Bas64 Conversions"
        ''' <summary>
        ''' It doesnt throw exception
        ''' </summary>
        ''' <param name="pBytesFile"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toBase64String(pBytesFile() As Byte) As String
            Try
                Return Convert.ToBase64String(pBytesFile)
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function
        ''' <summary>
        ''' It doesnt throw exception
        ''' </summary>
        ''' <param name="pImage"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toBase64String(pImage As Image) As String
            Return toBase64String(toByteArray(pImage))
        End Function


        ''' <summary>
        ''' It doesnt throw exception
        ''' </summary>
        ''' <param name="pBase64Data">Dont add the decoding data [data:image/gif;base64,]</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function fromBase64String(pBase64Data As String) As Byte()
            Try
                Dim p = Convert.FromBase64String(pBase64Data)
                Return p
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


#End Region



    End Class



End Namespace