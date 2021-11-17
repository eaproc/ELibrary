using CODERiT.Logger.v._3._5.Exceptions;

namespace ELibrary.Standard.Caching.Exceptions
{
    public class InvalidCodeException : EException
    {
        public InvalidCodeException(string FormTypeName, string ThreadAptName) : base(string.Format("Invalid Code Usage. Make sure the form [ {0} ] you are trying to load conforms to your Thread Apartment State [ {1} ].", FormTypeName, ThreadAptName))
        {
        }
    }
}