using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class WildCardValueMatcher : ValueMatcher
    {
        protected override int CompareStringToString(string s1, string s2)
        {
            if (s1.Contains("*") || s1.Contains("?"))
            {
                var regexPattern = Regex.Escape(s1);
                regexPattern = string.Format("^{0}$", regexPattern);
                regexPattern = regexPattern.Replace(@"\*", ".*");
                regexPattern = regexPattern.Replace(@"\?", ".");
                if (Regex.IsMatch(s2, regexPattern))
                {
                    return 0;
                }
            }
            return base.CompareStringToString(s1, s2);
        }
    }
}
