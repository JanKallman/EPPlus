using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public class NameValueProvider : INameValueProvider
    {
        private NameValueProvider()
        {

        }

        public static INameValueProvider Empty
        {
            get { return new NameValueProvider(); }
        }

        public bool IsNamedValue(string key, string worksheet)
        {
            return false;
        }

        public object GetNamedValue(string key)
        {
            return null;
        }

        public void Reload()
        {
            
        }
    }
}
