using System;
using System.ComponentModel;

namespace EPPlus.Utils
{
    internal static class EnumHelper
    {
        public static string GetDescriptionValue(Type enumType, object enumValue)
        {
            if (enumValue != null)
            {
                if (Nullable.GetUnderlyingType(enumType) != null)
                {
                    enumType = Nullable.GetUnderlyingType(enumType);
                }

                var fieldInfo = enumType.GetField(enumValue.ToString());

                var descriptionAttributes = fieldInfo.GetCustomAttributes(
                    typeof(DescriptionAttribute), false) as DescriptionAttribute[];

                if (descriptionAttributes == null) return enumValue.ToString();
                return (descriptionAttributes.Length > 0) ? descriptionAttributes[0].Description : enumValue.ToString();
            }

            return null;
        }
    }

}
