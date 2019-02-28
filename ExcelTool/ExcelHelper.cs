using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTool
{
    public static class ExcelHelper
    {
        public static T StringTo<T>(string input)
        {
            var type = typeof(T);
            if (type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                if (string.IsNullOrEmpty(input))
                {
                    return default(T);
                }

                type = Nullable.GetUnderlyingType(type);
            }
            return (T)Convert.ChangeType(input, type);
        }
    }
}
