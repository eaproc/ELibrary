Imports System.Runtime.Serialization
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO

Namespace Modules
    Public Module basObjectCloner

        <Extension()> _
        Public Function Clone(Of T)(source As T) As T

            If Not GetType(T).IsSerializable() Then Throw New ArgumentException("The type must be serializable.", "source")

            If (Object.ReferenceEquals(source, Nothing)) Then Return Nothing


            Dim formatter As IFormatter = New BinaryFormatter()
            Dim stream As Stream = New MemoryStream()
            Using (stream)

                formatter.Serialize(stream, source)
                stream.Seek(0, SeekOrigin.Begin)
                Return CType(formatter.Deserialize(stream), T)
            End Using

        End Function

    End Module

End Namespace