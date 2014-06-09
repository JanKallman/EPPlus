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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing
{
    internal static class DependencyChainFactory
    {
        internal static DependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            var depChain = new DependencyChain();
            foreach (var ws in wb.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    GetChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
                    GetWorksheetNames(ws, depChain, options);
                }
            }
            foreach (var name in wb.Names)
            {
                if (name.NameValue==null)
                {
                    GetChain(depChain, wb.FormulaParser.Lexer, name, options);
                }
            }
            return depChain;
        }

        internal static DependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            ws.CheckSheetType();
            var depChain = new DependencyChain();

            GetChain(depChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);

            GetWorksheetNames(ws, depChain, options);

            return depChain;
        }
        internal static DependencyChain Create(ExcelWorksheet ws, string Formula, ExcelCalculationOption options)
        {
            ws.CheckSheetType();
            var depChain = new DependencyChain();

            GetChain(depChain, ws.Workbook.FormulaParser.Lexer, ws, Formula, options);
            
            return depChain;
        }

        private static void GetWorksheetNames(ExcelWorksheet ws, DependencyChain depChain, ExcelCalculationOption options)
        {
            foreach (var name in ws.Names)
            {
                if (!string.IsNullOrEmpty(name.NameFormula))
                {
                    GetChain(depChain, ws.Workbook.FormulaParser.Lexer, name, options);
                }
            }
        }
        internal static DependencyChain Create(ExcelRangeBase range, ExcelCalculationOption options)
        {
            var depChain = new DependencyChain();

            GetChain(depChain, range.Worksheet.Workbook.FormulaParser.Lexer, range, options);

            return depChain;
        }
        private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelNamedRange name, ExcelCalculationOption options)
        {
            var ws = name.Worksheet;
            var id = ExcelCellBase.GetCellID(ws==null?0:ws.SheetID, name.Index, 0);
            if (!depChain.index.ContainsKey(id))
            {
                var f = new FormulaCell() { SheetID = ws == null ? 0 : ws.SheetID, Row = name.Index, Column = 0, Formula=name.NameFormula };
                if (!string.IsNullOrEmpty(f.Formula))
                {
                    f.Tokens = lexer.Tokenize(f.Formula, (ws==null ? null : ws.Name)).ToList();
                    if (ws == null)
                    {
                        name._workbook._formulaTokens.SetValue(name.Index, 0, f.Tokens);
                    }
                    else
                    {
                        ws._formulaTokens.SetValue(name.Index, 0, f.Tokens);
                    }
                    depChain.Add(f);
                    FollowChain(depChain, lexer,name._workbook, ws, f, options);
                }
            }
        }
        private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelWorksheet ws, string formula, ExcelCalculationOption options)
        {
            var f = new FormulaCell() { SheetID = ws.SheetID, Row = -1, Column = -1 };
            f.Formula = formula;
            if (!string.IsNullOrEmpty(f.Formula))
            {
                f.Tokens = lexer.Tokenize(f.Formula, ws.Name).ToList();
                depChain.Add(f);
                FollowChain(depChain, lexer, ws.Workbook, ws, f, options);
            }
        }

        private static void GetChain(DependencyChain depChain, ILexer lexer, ExcelRangeBase Range, ExcelCalculationOption options)
        {
            var ws = Range.Worksheet;
            var fs = new CellsStoreEnumerator<object>(ws._formulas, Range.Start.Row, Range.Start.Column, Range.End.Row, Range.End.Column);
            while (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellID(ws.SheetID, fs.Row, fs.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    var f = new FormulaCell() { SheetID = ws.SheetID, Row = fs.Row, Column = fs.Column };
                    if (fs.Value is int)
                    {
                        f.Formula = ws._sharedFormulas[(int)fs.Value].GetFormula(fs.Row, fs.Column, ws.Name);
                    }
                    else
                    {
                        f.Formula = fs.Value.ToString();
                    }
                    if (!string.IsNullOrEmpty(f.Formula))
                    {
                        f.Tokens = lexer.Tokenize(f.Formula, Range.Worksheet.Name).ToList();
                        ws._formulaTokens.SetValue(fs.Row, fs.Column, f.Tokens);
                        depChain.Add(f);
                        FollowChain(depChain, lexer, ws.Workbook, ws, f, options);
                    }
                }
            }
        }
        /// <summary>
        /// This method follows the calculation chain to get the order of the calculation
        /// Goto (!) is used internally to prevent stackoverflow on extremly larget dependency trees (that is, many recursive formulas).
        /// </summary>
        /// <param name="depChain">The dependency chain object</param>
        /// <param name="lexer">The formula tokenizer</param>
        /// <param name="wb">The workbook where the formula comes from</param>
        /// <param name="ws">The worksheet where the formula comes from</param>
        /// <param name="f">The cell function object</param>
        private static void FollowChain(DependencyChain depChain, ILexer lexer, ExcelWorkbook wb, ExcelWorksheet ws, FormulaCell f, ExcelCalculationOption options)
        {
            Stack<FormulaCell> stack = new Stack<FormulaCell>();
        iterateToken:
            while (f.tokenIx < f.Tokens.Count)
            {
                var t = f.Tokens[f.tokenIx];
                if (t.TokenType == TokenType.ExcelAddress)
                {
                    var adr = new ExcelFormulaAddress(t.Value);
                    if (adr.Table != null)
                    {
                        adr.SetRCFromTable(ws._package, new ExcelAddressBase(f.Row, f.Column, f.Row, f.Column));
                    }

                    if (adr.WorkSheet == null && adr.Collide(new ExcelAddressBase(f.Row, f.Column, f.Row, f.Column))!=ExcelAddressBase.eAddressCollition.No)
                    {
                        throw (new CircularReferenceException(string.Format("Circular Reference in cell {0}", ExcelAddressBase.GetAddress(f.Row, f.Column))));
                    }

                    if (adr._fromRow > 0 && adr._fromCol > 0)
                    {                        
                        if (string.IsNullOrEmpty(adr.WorkSheet))
                        {
                            if (f.ws == null)
                            {
                                f.ws = ws;
                            }
                            else if (f.ws.SheetID != f.SheetID)
                            {
                                f.ws = wb.Worksheets.GetBySheetID(f.SheetID);
                            }
                        }
                        else
                        {
                            f.ws = wb.Worksheets[adr.WorkSheet];
                        }

                        if (f.ws != null)
                        {
                            f.iterator = new CellsStoreEnumerator<object>(f.ws._formulas, adr.Start.Row, adr.Start.Column, adr.End.Row, adr.End.Column);
                            goto iterateCells;
                        }
                    }
                }
                else if (t.TokenType == TokenType.NameValue)
                {
                    string adrWb, adrWs, adrName;
                    ExcelNamedRange name;
                    ExcelAddressBase.SplitAddress(t.Value, out adrWb, out adrWs, out adrName, f.ws==null ? "" : f.ws.Name);
                    if (!string.IsNullOrEmpty(adrWs))
                    {
                        if (f.ws == null)
                        {
                            f.ws = wb.Worksheets[adrWs];
                        }
                        if(f.ws.Names.ContainsKey(t.Value))
                        {
                            name = f.ws.Names[adrName];
                        }
                        else if (wb.Names.ContainsKey(adrName))
                        {
                            name = wb.Names[adrName];
                        }
                        else
                        {
                            name = null;
                        }
                        if(name != null) f.ws = name.Worksheet;                        
                    }
                    else if (wb.Names.ContainsKey(adrName))
                    {
                        name = wb.Names[t.Value];
                        if (string.IsNullOrEmpty(adrWs))
                        {
                            f.ws = name.Worksheet;
                        }
                    }
                    else
                    {
                        name = null;
                    }

                    if (name != null)
                    {
                        if (string.IsNullOrEmpty(name.NameFormula))
                        {
                            f.iterator = new CellsStoreEnumerator<object>(f.ws._formulas, name.Start.Row, name.Start.Column, name.End.Row, name.End.Column);
                            goto iterateCells;
                        }
                        else
                        {
                            var id = ExcelAddressBase.GetCellID(name.LocalSheetId, name.Index, 0);

                            if (!depChain.index.ContainsKey(id))
                            {
                                var rf = new FormulaCell() { SheetID = name.LocalSheetId, Row = name.Index, Column = 0 };
                                rf.Formula = name.NameFormula;
                                rf.Tokens = lexer.Tokenize(rf.Formula, wb.Worksheets.GetBySheetID(name.LocalSheetId).Name).ToList();
                                depChain.Add(rf);
                                stack.Push(f);
                                f = rf;
                                goto iterateToken;
                            }
                            else
                            {
                                if (stack.Count > 0)
                                {
                                    //Check for circular references
                                    foreach (var par in stack)
                                    {
                                        if (ExcelAddressBase.GetCellID(par.SheetID, par.Row, par.Column) == id)
                                        {
                                            throw (new CircularReferenceException(string.Format("Circular Reference in name {0}", name.Name)));
                                        }
                                    }
                                }
                            }
                        }
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
                var v = f.iterator.Value;
                if (v == null || v.ToString().Trim() == "") continue;
                var id = ExcelAddressBase.GetCellID(f.ws.SheetID, f.iterator.Row, f.iterator.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    var rf = new FormulaCell() { SheetID = f.ws.SheetID, Row = f.iterator.Row, Column = f.iterator.Column };
                    if (f.iterator.Value is int)
                    {
                        rf.Formula = f.ws._sharedFormulas[(int)v].GetFormula(f.iterator.Row, f.iterator.Column, ws.Name);
                    }
                    else
                    {
                        rf.Formula = v.ToString();
                    }
                    rf.ws = f.ws;
                    rf.Tokens = lexer.Tokenize(rf.Formula, f.ws.Name).ToList();
                    ws._formulaTokens.SetValue(rf.Row, rf.Column, rf.Tokens);
                    depChain.Add(rf);
                    stack.Push(f);
                    f = rf;
                    goto iterateToken;
                }
                else
                {
                    if (stack.Count > 0)
                    {
                        //Check for circular references
                        foreach (var par in stack)
                        {
                            if (ExcelAddressBase.GetCellID(par.ws.SheetID, par.iterator.Row, par.iterator.Column) == id)
                            {
                                if (options.AllowCirculareReferences == false)
                                {
                                    throw (new CircularReferenceException(string.Format("Circular Reference in cell {0}!{1}", par.ws.Name, ExcelAddress.GetAddress(f.Row, f.Column))));
                                }
                                else
                                {
                                    f = stack.Pop();
                                    goto iterateCells;
                                }
                            }
                        }
                    }
                }
            }
            f.tokenIx++;
            goto iterateToken;
        }
    }
}
