using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
namespace EPPlus.Core.Compatibility
{
    internal class TypeCompat
    {
        public static bool IsPrimitive(object v)
        {
#if (Core)            
            return v.GetType().GetTypeInfo().IsPrimitive;
#else
            return v.GetType().IsPrimitive;
#endif
        }
        public static bool IsSubclassOf(Type t, Type c)
        {
#if (Core)            
            return t.GetTypeInfo().IsSubclassOf(c);
#else
            return t.IsSubclassOf(c);
#endif
        }

        internal static bool IsGenericType(Type t)
        {
#if (Core)            
            return t.GetTypeInfo().IsGenericType;
#else
            return t.IsGenericType;
#endif

        }
        public static object GetPropertyValue(object v, string name)
        {
#if (Core)
            return v.GetType().GetTypeInfo().GetProperty(name).GetValue(v, null);
#else
            return v.GetType().GetProperty(name).GetValue(v, null);
#endif
        }
    }
}
