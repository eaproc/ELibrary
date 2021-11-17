using System.Collections.Generic;
using ELibrary.Standard.Modules;

namespace ELibrary.Standard.Objects
{
    public class IEqualityComparerIgnoreCase : IEqualityComparer<string>
    {

        #region IEquality Comparer - Ignore Case
        bool IEqualityComparer<string>.Equals(string x, string y)
        {
            return x.equalsIgnoreCase(y);
        }

        public bool Equals1(string x, string y) => ((IEqualityComparer<string>)this).Equals(x, y);

        int IEqualityComparer<string>.GetHashCode(string obj)
        {
            return obj.GetHashCode();
        }

        public int GetHashCode1(string obj) => ((IEqualityComparer<string>)this).GetHashCode(obj);

        #endregion

    }
}