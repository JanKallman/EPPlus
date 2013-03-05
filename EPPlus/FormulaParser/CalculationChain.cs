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
 * Jan Källman                      Added                       2012-03-04  
 *******************************************************************************/
using ExcelFormulaParser.Engine.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParser
{
    internal static class CalculationChain
    {
        internal static DependencyChain GetChain(ExcelWorkbook wb)
        {
            var depChain = new DependencyChain();
            var d = new Dictionary<string, object>();
            var sct = new SourceCodeTokenizer(d, d);
            foreach (var ws in wb.Worksheets)
            {
                GetChain(depChain, sct, ws.Cells);
            }
            return depChain;
        }
        internal static DependencyChain GetChain(ExcelWorksheet ws)
        {
            var depChain = new DependencyChain();
            var d = new Dictionary<string, object>();
            var sct = new SourceCodeTokenizer(d, d);

            GetChain(depChain, sct, ws.Cells);

            return depChain;
        }
        internal static DependencyChain GetChain(ExcelRangeBase range)
        {
            var depChain = new DependencyChain();
            var d = new Dictionary<string, object>();
            var sct = new SourceCodeTokenizer(d, d);

            GetChain(depChain, sct, range);

            return depChain;
        }
        private static void GetChain(DependencyChain depChain, SourceCodeTokenizer sct, ExcelRangeBase Range)
        {
            var ws = Range.Worksheet;
            var fs = new CellsStoreEnumerator<object>(ws._formulas, Range.Start.Row, Range.Start.Column, Range.End.Row, Range.End.Column);
            while (fs.Next())
            {
                var id = ExcelCellBase.GetCellID(ws.SheetID, fs.Row, fs.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    var f = new FormulaCell() { SheetID = ws.SheetID, Row = fs.Row, Column = fs.Column };
                    if (fs.Value is int)
                    {
                        f.Formula = ws._sharedFormulas[(int)fs.Value].GetFormula(fs.Row, fs.Column);
                    }
                    else
                    {
                        f.Formula = fs.Value.ToString();
                    }
                    f.Tokens = sct.Tokenize(f.Formula).ToList();
                    depChain.Add(f);
                    FollowChain(depChain, sct, ws, f);
                }
            }
        }
        /// <summary>
        /// This method follows the calculation chain to get the order of the calculation
        /// Goto (!) is used internally to prevent stackoverflow on extremly larget dependency trees (that is many recursive formulas).
        /// </summary>
        /// <param name="depChain">The dependency chain object</param>
        /// <param name="sct">The formula tokenizer</param>
        /// <param name="ws">The worksheet where the formula comes from</param>
        /// <param name="f">The cell function obleject</param>
        private static void FollowChain(DependencyChain depChain, SourceCodeTokenizer sct, ExcelWorksheet ws, FormulaCell f)
        {
            Stack<FormulaCell> stack = new Stack<FormulaCell>();
        iterateToken:
            while (f.tokenIx < f.Tokens.Count)
            {
                var t = f.Tokens[f.tokenIx];
                if (t.TokenType == TokenType.ExcelAddress)
                {
                    var adr = new ExcelFormulaAddress(t.Value);
                    if (string.IsNullOrEmpty(adr.WorkSheet))
                    {
                        f.ws = ws;
                    }
                    else
                    {
                        f.ws = ws.Workbook.Worksheets[adr.WorkSheet];
                    }
                    if (f.ws != null)
                    {
                        f.iterator = new CellsStoreEnumerator<object>(f.ws._formulas, adr.Start.Row, adr.Start.Column, adr.End.Row, adr.End.Column);
                        goto iterateCells;
                    }
                }
                f.tokenIx++;
            }
            depChain.CalcOrder.Add(f.Index);
            if (stack.Count > 0)
            {
                f = stack.Pop();
                goto iterateCells;
            }
            return;
        iterateCells:

            while (f.iterator.Next())
            {
                var id = ExcelAddressBase.GetCellID(f.ws.SheetID, f.iterator.Row, f.iterator.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    var rf = new FormulaCell() { SheetID = f.ws.SheetID, Row = f.iterator.Row, Column = f.iterator.Column };
                    if (f.iterator.Value is int)
                    {
                        rf.Formula = f.ws._sharedFormulas[(int)f.iterator.Value].GetFormula(f.iterator.Row, f.iterator.Column);
                    }
                    else
                    {
                        rf.Formula = f.iterator.Value.ToString();
                    }
                    rf.Tokens = sct.Tokenize(rf.Formula).ToList();
                    depChain.Add(rf);
                    stack.Push(f);
                    f = rf;
                    goto iterateToken;
                }
                else if (stack.Count > 0)
                {
                    //Check for circular references
                    foreach (var par in stack)
                    {
                        if (ExcelAddressBase.GetCellID(par.ws.SheetID, par.iterator.Row, par.iterator.Column) == id)
                        {
                            throw (new ExcelFormulaParser.Engine.Exceptions.CircularReferenceException(string.Format("Circular Reference in cell {0}!{1}", par.ws.Name, ExcelAddress.GetAddress(f.Row, f.Column))));
                        }
                    }
                }
            }
            f.tokenIx++;
            goto iterateToken;
        }
    }
    internal class DependencyChain
    {
        internal List<FormulaCell> list = new List<FormulaCell>();
        internal Dictionary<ulong, int> index = new Dictionary<ulong, int>();
        internal List<int> CalcOrder = new List<int>();
        internal void Add(FormulaCell f)
        {
            list.Add(f);
            f.Index = list.Count - 1;
            index.Add(ExcelCellBase.GetCellID(f.SheetID, f.Row, f.Column), f.Index);
        }
    }
    internal class FormulaCell
    {
        internal int Index { get; set; }
        internal int SheetID { get; set; }
        internal int Row { get; set; }
        internal int Column { get; set; }
        internal string Formula { get; set; }
        internal List<Token> Tokens { get; set; }

        internal int tokenIx = 0;
        internal int addressIx = 0;
        internal CellsStoreEnumerator<object> iterator;
        internal ExcelWorksheet ws;
    }
}
