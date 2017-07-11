using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Compatibility
{
    // System.Security.SecurityElement
#if Core
    public class SecurityElement
        {
        private static readonly char[] _chars = new char[]
        {
            '<',
            '>',
            '"',
            '\'',
            '&'
        };

        public static string Escape(string str)
        {
            if (str == null)
            {
                return null;
            }
            var ix = str.IndexOfAny(_chars);
            if (ix < 0) return str;
            var sb = new StringBuilder(str.Substring(0,ix));
            for (int i=ix;i < str.Length;i++)
            {
                switch(str[i])
                {
                    case '<':
                        sb.Append("&lt");
                        break;
                    case '>':
                        sb.Append("&gt");
                        break;
                    case '"':
                        sb.Append("&quot;");
                        break;
                    case '\'':
                        sb.Append("&apos;");
                        break;
                    case '&':
                        sb.Append("&amp;");
                        break;
                    default:
                        sb.Append(str[i]);
                        break;
                }
            }
            return sb.ToString();
        }
    }
#endif
}
