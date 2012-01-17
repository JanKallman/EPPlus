/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-08
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaTime : ExcelDataValidationFormulaValue<ExcelTime>, IExcelDataValidationFormulaTime
    {
        public ExcelDataValidationFormulaTime(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
            : base(namespaceManager, topNode, formulaPath)
        {
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                decimal time = default(decimal);
                if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out time))
                {
                    Value = new ExcelTime(time);
                }
                else
                {
                    Value = new ExcelTime();
                    ExcelFormula = value;
                }
            }
            else
            {
                Value = new ExcelTime();
            }
            Value.TimeChanged += new EventHandler(Value_TimeChanged);
        }

        void Value_TimeChanged(object sender, EventArgs e)
        {
            SetXmlNodeString(FormulaPath, Value.ToExcelString());
        }

        protected override string GetValueAsString()
        {
            if (State == FormulaState.Value)
            {
                return Value.ToExcelString();
            }
            return string.Empty;
        }

        internal override void ResetValue()
        {
            Value = new ExcelTime();
        }
    }
}
