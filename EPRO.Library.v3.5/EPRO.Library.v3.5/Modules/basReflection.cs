using System.Collections.Generic;
using System.Linq;

namespace ELibrary.Standard.Modules
{
    public static class basReflection
    {
        public static object getAnonymousClassPropValue(this object pClass, string pPropertyName, object[] pIndex = null)
        {
            if (pClass is null)
                return null;
            var myType = pClass.GetType();
            IList<System.Reflection.PropertyInfo> props = myType.GetProperties();
            IList<System.Reflection.FieldInfo> fields = myType.GetFields();
            if (props.Count > 0)
            {
                var prop = props.Where(x => (x.Name ?? "") == (pPropertyName ?? "")).FirstOrDefault();
                if (prop is object)
                    return prop.GetValue(pClass, pIndex);
            }
            else if (fields.Count > 0)
            {
                var field = fields.Where(x => (x.Name ?? "") == (pPropertyName ?? "")).FirstOrDefault();
                if (field is object)
                    return field.GetValue(pClass);
            }

            return null;
        }

        public static System.Reflection.MethodInfo getAnonymousClassMethod(this object pClass, string pMethodName)
        {
            if (pClass is null)
                return null;
            var myType = pClass.GetType();
            IList<System.Reflection.MethodInfo> meths = myType.GetMethods();
            if (meths.Count > 0)
            {
                var meth = meths.Where(x => (x.Name ?? "") == (pMethodName ?? "")).FirstOrDefault();
                if (meth is object)
                    return meth;
            }

            return null;
        }
    }
}