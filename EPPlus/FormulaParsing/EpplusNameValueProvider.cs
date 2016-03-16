using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusNameValueProvider : INameValueProvider
    {
        private ExcelDataProvider _excelDataProvider;
        private ExcelNamedRangeCollection _values;

        public EpplusNameValueProvider(ExcelDataProvider excelDataProvider)
        {
            _excelDataProvider = excelDataProvider;
            _values = _excelDataProvider.GetWorkbookNameValues();
        }

        public virtual bool IsNamedValue(string key, string ws)
        {
            if(ws!=null)
            {
                var wsNames = _excelDataProvider.GetWorksheetNames(ws);
                if(wsNames!=null && wsNames.ContainsKey(key))
                {
                    return true;
                }
            }
            return _values != null && _values.ContainsKey(key);
        }

        public virtual object GetNamedValue(string key)
        {
            return _values[key];
        }

        public virtual void Reload()
        {
            _values = _excelDataProvider.GetWorkbookNameValues();
        }
    }
}
