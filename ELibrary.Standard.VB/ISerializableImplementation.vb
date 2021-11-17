Imports System.Runtime.Serialization

Namespace Objects
    Public Interface ISerializableImplementation
        Inherits ISerializable

        ''' <summary>
        ''' Mark the class [ Serializable() _ ]. Also implement Protected Sub New(info,context) call Deserialize(info,context) 
        ''' info.getValue() as you used info.addValue() under serialization
        ''' </summary>
        ''' <param name="info"></param>
        ''' <param name="context"></param>
        ''' <remarks></remarks>
        Sub Deserialize(ByVal info As SerializationInfo, ByVal context As StreamingContext)


    End Interface

End Namespace