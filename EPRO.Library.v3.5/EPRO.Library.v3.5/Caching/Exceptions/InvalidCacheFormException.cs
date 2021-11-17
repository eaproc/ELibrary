using CODERiT.Logger.v._3._5.Exceptions;

namespace ELibrary.Standard.Caching.Exceptions
{
    public class InvalidCacheFormException : EException
    {
        public InvalidCacheFormException() : base("Invalid Cache Form Type")
        {
        }
    }
}