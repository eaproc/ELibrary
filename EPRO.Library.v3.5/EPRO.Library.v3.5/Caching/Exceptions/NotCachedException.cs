using CODERiT.Logger.v._3._5.Exceptions;

namespace ELibrary.Standard.Caching.Exceptions
{
    public class NotCachedException : EException
    {
        public NotCachedException(string FormTypeName) : base(string.Format("Invalid Code Usage. The requested Form [ {0} ] is NOT cached!!!", FormTypeName))
        {
        }
    }
}