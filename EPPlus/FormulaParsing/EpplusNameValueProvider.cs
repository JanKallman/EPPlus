/* Copyright (C) 2011  Jan Källman
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/
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
