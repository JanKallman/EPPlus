using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public static class ExcelAddressUtil
    {
        static char[] SheetNameInvalidChars = new char[] { '?', ':', '*', '/', '\\' };
        public static bool IsValidAddress(string token)
        {
            int ix;
            if (token[0] == '\'')
            {
                ix = token.LastIndexOf('\'');
                if (ix > 0 && ix < token.Length - 1 && token[ix + 1] == '!')
                {
                    if (token.IndexOfAny(SheetNameInvalidChars, 1, ix - 1) > 0)
                    {
                        return false;
                    }
                    token = token.Substring(ix + 2);
                }
                else
                {
                    return false;
                }
            }
            else if ((ix = token.IndexOf('!')) > 1)
            {
                if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
                {
                    return false;
                }
                token = token.Substring(token.IndexOf('!') + 1);
            }
            return OfficeOpenXml.ExcelAddress.IsValidAddress(token);
        }
    }
}
