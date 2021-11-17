using System.Runtime.Serialization;

namespace ELibrary.Standard.Objects
{
    public interface ISerializableImplementation : ISerializable
    {

        /// <summary>
        /// Mark the class [ Serializable() _ ]. Also implement Protected Sub New(info,context) call Deserialize(info,context) 
        /// info.getValue() as you used info.addValue() under serialization
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        /// <remarks></remarks>
        void Deserialize(SerializationInfo info, StreamingContext context);
    }
}