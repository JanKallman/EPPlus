﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman                      Added                       2012-03-04  
 *******************************************************************************/

using System.Threading;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml
{
    public static class CalculationExtension
    {

        public static void Calculate(this ExcelWorkbook workbook)
        {
            Calculate(workbook, new ExcelCalculationOption(){AllowCirculareReferences=false});
        }
        public static void Calculate(this ExcelWorkbook workbook, ExcelCalculationOption options)
        {
            Init(workbook);

            var dc = DependencyChainFactory.Create(workbook, options);
            workbook.FormulaParser.InitNewCalc();
            if (workbook.FormulaParser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                workbook.FormulaParser.Logger.Log(msg);
            }

            //TODO: Remove when tests are done. Outputs the dc to a text file. 
            //var fileDc = new System.IO.StreamWriter("c:\\temp\\dc.txt");
                        
            //for (int i = 0; i < dc.list.Count; i++)
            //{
            //    fileDc.WriteLine(i.ToString() + "," + dc.list[i].Column.ToString() + "," + dc.list[i].Row.ToString() + "," + (dc.list[i].ws==null ? "" : dc.list[i].ws.Name) + "," + dc.list[i].Formula);
            //}
            //fileDc.Close();
            //fileDc = new System.IO.StreamWriter("c:\\temp\\dcorder.txt");
            //for (int i = 0; i < dc.CalcOrder.Count; i++)
            //{
            //    fileDc.WriteLine(dc.CalcOrder[i].ToString());
            //}
            //fileDc.Close();
            //fileDc = null;

            //TODO: Add calculation here

            CalcChain(workbook, workbook.FormulaParser, dc);

            //workbook._isCalculated = true;
        }
        public static void Calculate(this ExcelWorksheet worksheet)
        {
            Calculate(worksheet, new ExcelCalculationOption());
        }
        public static void Calculate(this ExcelWorksheet worksheet, ExcelCalculationOption options)
        {
            Init(worksheet.Workbook);
            //worksheet.Workbook._formulaParser = null; TODO:Cant reset. Don't work with userdefined or overrided worksheet functions            
            var dc = DependencyChainFactory.Create(worksheet, options);
            var parser = worksheet.Workbook.FormulaParser;
            parser.InitNewCalc();
            if (parser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                parser.Logger.Log(msg);
            }
            CalcChain(worksheet.Workbook, parser, dc);
        }
        public static void Calculate(this ExcelRangeBase range)
        {
            Calculate(range, new ExcelCalculationOption());
        }
        public static void Calculate(this ExcelRangeBase range, ExcelCalculationOption options)
        {
            Init(range._workbook);
            var parser = range._workbook.FormulaParser;
            parser.InitNewCalc();
            var dc = DependencyChainFactory.Create(range, options);
            CalcChain(range._workbook, parser, dc);
        }
        public static object Calculate(this ExcelWorksheet worksheet, string Formula)
        {
            return Calculate(worksheet, Formula, new ExcelCalculationOption());
        }
        public static object Calculate(this ExcelWorksheet worksheet, string Formula, ExcelCalculationOption options)
        {
            try
            {
                worksheet.CheckSheetType();
                if(string.IsNullOrEmpty(Formula.Trim())) return null;
                Init(worksheet.Workbook);
                var parser = worksheet.Workbook.FormulaParser;
                parser.InitNewCalc();
                if (Formula[0] == '=') Formula = Formula.Substring(1); //Remove any starting equal sign
                var dc = DependencyChainFactory.Create(worksheet, Formula, options);
                var f = dc.list[0];
                dc.CalcOrder.RemoveAt(dc.CalcOrder.Count - 1);

                CalcChain(worksheet.Workbook, parser, dc);

                return parser.ParseCell(f.Tokens, worksheet.Name, -1, -1);
            }
            catch (Exception ex)
            {
                return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
            }
        }
        private static void CalcChain(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc)
        {
            var debug = parser.Logger != null;
            foreach (var ix in dc.CalcOrder)
            {
                var item = dc.list[ix];
                try
                {
                    var ws = wb.Worksheets.GetBySheetID(item.SheetID);
                    var v = parser.ParseCell(item.Tokens, ws == null ? "" : ws.Name, item.Row, item.Column);
                    SetValue(wb, item, v);
                    if (debug)
                    {
                        parser.Logger.LogCellCounted();
                    }
                    Thread.Sleep(0);
                }
                catch (FormatException fe)
                {
                    throw (fe);
                }
                catch(Exception e)
                {
                    var error = ExcelErrorValue.Parse(ExcelErrorValue.Values.Value);
                    SetValue(wb, item, error);
                }
            }
        }
        private static void Init(ExcelWorkbook workbook)
        {
            workbook._formulaTokens = new CellStore<List<Token>>();;
            foreach (var ws in workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    if (ws._formulaTokens != null)
                    {
                        ws._formulaTokens.Dispose();
                    }
                    ws._formulaTokens = new CellStore<List<Token>>();
                }
            }
        }

        private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
        {
            if (item.Column == 0)
            {
                if (item.SheetID <= 0)
                {
                    workbook.Names[item.Row].NameValue = v;
                }
                else
                {
                    var sh = workbook.Worksheets.GetBySheetID(item.SheetID);
                    sh.Names[item.Row].NameValue = v;
                }
            }
            else
            {
                var sheet = workbook.Worksheets.GetBySheetID(item.SheetID);
                sheet.SetValueInner(item.Row, item.Column, v);
            }
        }
    }
}
