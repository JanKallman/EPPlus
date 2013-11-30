using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class BuiltInFunctions : FunctionsModule
    {
        public BuiltInFunctions()
        {
            // Text
            Functions["text"] = new CStr();
            Functions["len"] = new Len();
            Functions["lower"] = new Lower();
            Functions["upper"] = new Upper();
            Functions["left"] = new Left();
            Functions["right"] = new Right();
            Functions["mid"] = new Mid();
            Functions["replace"] = new Replace();
            Functions["substitute"] = new Substitute();
            Functions["concatenate"] = new Concatenate();
            // Numbers
            Functions["int"] = new CInt();
            // Math
            Functions["cos"] = new Cos();
            Functions["cosh"] = new Cosh();
            Functions["power"] = new Power();
            Functions["sqrt"] = new Sqrt();
            Functions["sqrtpi"] = new SqrtPi();
            Functions["pi"] = new Pi();
            Functions["product"] = new Product();
            Functions["ceiling"] = new Ceiling();
            Functions["count"] = new Count();
            Functions["counta"] = new CountA();
            Functions["floor"] = new Floor();
            Functions["sin"] = new Sin();
            Functions["sinh"] = new Sinh();
            Functions["sum"] = new Sum();
            Functions["sumif"] = new SumIf();
            Functions["sumproduct"] = new SumProduct();
            Functions["stdev"] = new Stdev();
            Functions["stdevp"] = new StdevP();
            Functions["subtotal"] = new Subtotal();
            Functions["exp"] = new Exp();
            Functions["log"] = new Log();
            Functions["log10"] = new Log10();
            Functions["max"] = new Max();
            Functions["maxa"] = new Maxa();
            Functions["min"] = new Min();
            Functions["mod"] = new Mod();
            Functions["average"] = new Average();
            Functions["round"] = new Round();
            Functions["rand"] = new Rand();
            Functions["randbetween"] = new RandBetween();
            Functions["tan"] = new Tan();
            Functions["tanh"] = new Tanh();
            Functions["atan"] = new Atan();
            Functions["atan2"] = new Atan2();
            Functions["var"] = new Var();
            Functions["varp"] = new VarP();
            // Information
            Functions["isblank"] = new IsBlank();
            Functions["isnumber"] = new IsNumber();
            Functions["istext"] = new IsText();
            Functions["iserror"] = new IsError();
            // Logical
            Functions["if"] = new If();
            Functions["not"] = new Not();
            Functions["and"] = new And();
            Functions["or"] = new Or();
            Functions["true"] = new True();
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
        }
    }
}
