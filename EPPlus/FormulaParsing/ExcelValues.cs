using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    public enum eErrorType
    {
        Div0,
        NA,
        Name,
        Null,
        Num,
        Ref,
        Value
    }
    public class ExcelErrorValue
    {
        public static class Values
        {
            public const string Div0 = "#DIV/0!";
            public const string NA = "#N/A";
            public const string Name = "#NAME?";
            public const string Null = "#NULL!";
            public const string Num = "#NUM!";
            public const string Ref = "#REF!";
            public const string Value = "#VALUES!";
        }
        internal static ExcelErrorValue Create(eErrorType errorType)
        {
            return new ExcelErrorValue(errorType);
        }

        private ExcelErrorValue(eErrorType type)
        {
            Type=type; 
        }
        public eErrorType Type { get; private set; }
        public override string ToString()
        {
            switch(Type)
            {
                case eErrorType.Div0:
                    return Values.Div0;
                case eErrorType.NA:
                    return Values.NA;
                case eErrorType.Name:
                    return Values.Name;
                case eErrorType.Null:
                    return Values.Null;
                case eErrorType.Num:
                    return Values.Num;
                case eErrorType.Ref:
                    return Values.Ref;
                case eErrorType.Value:
                    return Values.Value;
                default:
                    throw(new ArgumentException("Invalid errortype"));
            }
        }
        public static ExcelErrorValue operator +(object v1, ExcelErrorValue v2)
        {
            return v2;
        }
        public static ExcelErrorValue operator +(ExcelErrorValue v1, ExcelErrorValue v2)
        {
            return v1;
        }
    }
}
