using System.Collections.Generic;

namespace ELibrary.Standard.Objects
{
    public class LikeIgnoreCaseComparer : IEqualityComparer<string>
    {
        bool IEqualityComparer<string>.Equals(string x, string y)
        {
            if (x is null)
                return false;
            return x.ToLower().IndexOf(y.ToLower()) >= 0;
        }

        public bool Equals1(string x, string y) => ((IEqualityComparer<string>)this).Equals(x, y);

        int IEqualityComparer<string>.GetHashCode(string obj)
        {
            return obj.GetHashCode();
        }

        public int GetHashCode1(string obj) => ((IEqualityComparer<string>)this).GetHashCode(obj);
    }
}