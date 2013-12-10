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
        internal ExcelErrorValue(eErrorType type)
        {
            Type=type; 
        }
        public eErrorType Type { get; private set; }
        public override string ToString()
        {
            switch(Type)
            {
                case eErrorType.Div0:
                    return "#DIV/0!";
                case eErrorType.NA:
                    return "#N/A";
                case eErrorType.Name:
                    return "#NAME?";
                case eErrorType.Null:
                    return "#NULL!";
                case eErrorType.Num:
                    return "#NUM!";
                case eErrorType.Ref:
                    return "#REF!";
                case eErrorType.Value:
                    return "#VALUE!";
                default:
                    throw(new ArgumentException("Invalid errortype"));
            }
        }
    }
}
