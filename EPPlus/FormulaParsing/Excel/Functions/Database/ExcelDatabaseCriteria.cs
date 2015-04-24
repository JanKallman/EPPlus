/* Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

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
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2015-04-06
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class ExcelDatabaseCriteria
    {
        private readonly ExcelDataProvider _dataProvider;
        private readonly int _fromCol;
        private readonly int _toCol;
        private readonly string _worksheet;
        private readonly int _fieldRow;
        private readonly Dictionary<ExcelDatabaseCriteriaField, object> _criterias = new Dictionary<ExcelDatabaseCriteriaField, object>(); 

        public ExcelDatabaseCriteria(ExcelDataProvider dataProvider, string range)
        {
            _dataProvider = dataProvider;
            var address = new ExcelAddressBase(range);
            _fromCol = address._fromCol;
            _toCol = address._toCol;
            _worksheet = address.WorkSheet;
            _fieldRow = address._fromRow;
            Initialize();
        }

        private void Initialize()
        {
            var fo = 1;
            for (var x = _fromCol; x <= _toCol; x++)
            {
                var fieldObj = _dataProvider.GetCellValue(_worksheet, _fieldRow, x);
                var val = _dataProvider.GetCellValue(_worksheet, _fieldRow + 1, x);
                if (fieldObj != null && val != null)
                {
                    if(fieldObj is string)
                    { 
                        var field = new ExcelDatabaseCriteriaField(fieldObj.ToString().ToLower(CultureInfo.InvariantCulture));
                        _criterias.Add(field, val);
                    }
                    else if (ConvertUtil.IsNumeric(fieldObj))
                    {
                        var field = new ExcelDatabaseCriteriaField((int) fieldObj);
                        _criterias.Add(field, val);
                    }

                }
            }
        }

        public virtual IDictionary<ExcelDatabaseCriteriaField, object> Items
        {
            get { return _criterias; }
        }
    }
}
