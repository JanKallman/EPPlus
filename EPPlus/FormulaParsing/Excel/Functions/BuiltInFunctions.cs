﻿/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class BuiltInFunctions : FunctionsModule
    {
        public BuiltInFunctions()
        {
            // Text
            Functions["len"] = new Len();
            Functions["lower"] = new Lower();
            Functions["upper"] = new Upper();
            Functions["left"] = new Left();
            Functions["right"] = new Right();
            Functions["mid"] = new Mid();
            Functions["replace"] = new Replace();
            Functions["rept"] = new Rept();
            Functions["substitute"] = new Substitute();
            Functions["concatenate"] = new Concatenate();
            Functions["char"] = new CharFunction();
            Functions["exact"] = new Exact();
            Functions["find"] = new Find();
            Functions["fixed"] = new Fixed();
            Functions["proper"] = new Proper();
            Functions["search"] = new Search();
            Functions["text"] = new Text.Text();
            Functions["t"] = new T();
            Functions["hyperlink"] = new Hyperlink();
            Functions["value"] = new Value();
            // Numbers
            Functions["int"] = new CInt();
            // Math
            Functions["abs"] = new Abs();
            Functions["asin"] = new Asin();
            Functions["asinh"] = new Asinh();
            Functions["cos"] = new Cos();
            Functions["cosh"] = new Cosh();
            Functions["power"] = new Power();
            Functions["sign"] = new Sign();
            Functions["sqrt"] = new Sqrt();
            Functions["sqrtpi"] = new SqrtPi();
            Functions["pi"] = new Pi();
            Functions["product"] = new Product();
            Functions["ceiling"] = new Ceiling();
            Functions["count"] = new Count();
            Functions["counta"] = new CountA();
            Functions["countblank"] = new CountBlank();
            Functions["countif"] = new CountIf();
            Functions["countifs"] = new CountIfs();
            Functions["fact"] = new Fact();
            Functions["floor"] = new Floor();
            Functions["sin"] = new Sin();
            Functions["sinh"] = new Sinh();
            Functions["sum"] = new Sum();
            Functions["sumif"] = new SumIf();
            Functions["sumifs"] = new SumIfs();
            Functions["sumproduct"] = new SumProduct();
            Functions["sumsq"] = new Sumsq();
            Functions["stdev"] = new Stdev();
            Functions["stdevp"] = new StdevP();
            Functions["stdev.s"] = new Stdev();
            Functions["stdev.p"] = new StdevP();
            Functions["subtotal"] = new Subtotal();
            Functions["exp"] = new Exp();
            Functions["log"] = new Log();
            Functions["log10"] = new Log10();
            Functions["ln"] = new Ln();
            Functions["max"] = new Max();
            Functions["maxa"] = new Maxa();
            Functions["median"] = new Median();
            Functions["min"] = new Min();
            Functions["mina"] = new Mina();
            Functions["mod"] = new Mod();
            Functions["average"] = new Average();
            Functions["averagea"] = new AverageA();
            Functions["averageif"] = new AverageIf();
            Functions["averageifs"] = new AverageIfs();
            Functions["round"] = new Round();
            Functions["rounddown"] = new Rounddown();
            Functions["roundup"] = new Roundup();
            Functions["rand"] = new Rand();
            Functions["randbetween"] = new RandBetween();
            Functions["rank"] = new Rank();
            Functions["rank.eq"] = new Rank();
            Functions["rank.avg"] = new Rank(true);
            Functions["quotient"] = new Quotient();
            Functions["trunc"] = new Trunc();
            Functions["tan"] = new Tan();
            Functions["tanh"] = new Tanh();
            Functions["atan"] = new Atan();
            Functions["atan2"] = new Atan2();
            Functions["atanh"] = new Atanh();
            Functions["acos"] = new Acos();
            Functions["acosh"] = new Acosh();
            Functions["var"] = new Var();
            Functions["varp"] = new VarP();
            Functions["large"] = new Large();
            Functions["small"] = new Small();
            Functions["degrees"] = new Degrees();
            // Information
            Functions["isblank"] = new IsBlank();
            Functions["isnumber"] = new IsNumber();
            Functions["istext"] = new IsText();
            Functions["isnontext"] = new IsNonText();
            Functions["iserror"] = new IsError();
            Functions["iserr"] = new IsErr();
            Functions["error.type"] = new ErrorType();
            Functions["iseven"] = new IsEven();
            Functions["isodd"] = new IsOdd();
            Functions["islogical"] = new IsLogical();
            Functions["isna"] = new IsNa();
            Functions["na"] = new Na();
            Functions["n"] = new N();
            // Logical
            Functions["if"] = new If();
            Functions["iferror"] = new IfError();
            Functions["ifna"] = new IfNa();
            Functions["not"] = new Not();
            Functions["and"] = new And();
            Functions["or"] = new Or();
            Functions["true"] = new True();
            Functions["false"] = new False();
            // Reference and lookup
            Functions["address"] = new Address();
            Functions["hlookup"] = new HLookup();
            Functions["vlookup"] = new VLookup();
            Functions["lookup"] = new Lookup();
            Functions["match"] = new Match();
            Functions["row"] = new Row();
            Functions["rows"] = new Rows();
            Functions["column"] = new Column();
            Functions["columns"] = new Columns();
            Functions["choose"] = new Choose();
            Functions["index"] = new RefAndLookup.Index();
            Functions["indirect"] = new Indirect();
            Functions["offset"] = new Offset();
            // Date
            Functions["date"] = new Date();
            Functions["today"] = new Today();
            Functions["now"] = new Now();
            Functions["day"] = new Day();
            Functions["month"] = new Month();
            Functions["year"] = new Year();
            Functions["time"] = new Time();
            Functions["hour"] = new Hour();
            Functions["minute"] = new Minute();
            Functions["second"] = new Second();
            Functions["weeknum"] = new Weeknum();
            Functions["weekday"] = new Weekday();
            Functions["days360"] = new Days360();
            Functions["yearfrac"] = new Yearfrac();
            Functions["edate"] = new Edate();
            Functions["eomonth"] = new Eomonth();
            Functions["isoweeknum"] = new IsoWeekNum();
            Functions["workday"] = new Workday();
            Functions["networkdays"] = new Networkdays();
            Functions["networkdays.intl"] = new NetworkdaysIntl();
            Functions["datevalue"] = new DateValue();
            Functions["timevalue"] = new TimeValue();
            // Database
            Functions["dget"] = new Dget();
            Functions["dcount"] = new Dcount();
            Functions["dcounta"] = new DcountA();
            Functions["dmax"] = new Dmax();
            Functions["dmin"] = new Dmin();
            Functions["dsum"] = new Dsum();
            Functions["daverage"] = new Daverage();
            Functions["dvar"] = new Dvar();
            Functions["dvarp"] = new Dvarp();
            //Finance
            Functions["pmt"] = new Pmt();
        }
    }
}
