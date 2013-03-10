using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class IndexToAddressTranslator
    {
        public IndexToAddressTranslator(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ExcelReferenceType.AbsoluteRowAndColumn)
        {

        }

        public IndexToAddressTranslator(ExcelDataProvider excelDataProvider, ExcelReferenceType referenceType)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _excelReferenceType = referenceType;
        }

        const int MaxAlphabetIndex = 25;
        const int NLettersInAlphabet = 26;
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ExcelReferenceType _excelReferenceType;

        public string ToAddress(int col, int row)
        {
            if (col <= MaxAlphabetIndex)
            {
                return string.Concat(GetColumn(IntToChar(col)), GetRowNumber(row + 1));
            }
            else if (col < (Math.Pow(NLettersInAlphabet, 2) + NLettersInAlphabet))
            {
                var firstChar = col / NLettersInAlphabet - 1;
                var secondChar = col % NLettersInAlphabet;
                return string.Concat(GetColumn(IntToChar(firstChar), IntToChar(secondChar)), GetRowNumber(row + 1));
            }
            else if(col < (Math.Pow(NLettersInAlphabet, 3) + NLettersInAlphabet))
            {
                var x = NLettersInAlphabet * NLettersInAlphabet;
                var rest = col - x;
                var firstChar = col / x - 1;
                var secondChar = rest / NLettersInAlphabet - 1;
                var thirdChar = rest % NLettersInAlphabet;
                return string.Concat(GetColumn(IntToChar(firstChar), IntToChar(secondChar), IntToChar(thirdChar)), GetRowNumber(row + 1));
            }
            throw new InvalidOperationException("ExcelFormulaParser does not the supplied number of columns " + col);
        }

        private string GetColumn(params char[] chars)
        {
            var retVal = new StringBuilder().Append(chars).ToString();
            switch (_excelReferenceType)
            {
                case ExcelReferenceType.AbsoluteRowAndColumn:
                case ExcelReferenceType.RelativeRowAbsolutColumn:
                    return "$" + retVal;
                default:
                    return retVal;
            }
        }

        private char IntToChar(int i)
        {
            return (char)(i + 65);
        }

        private string GetRowNumber(int rowNo)
        {
            var retVal = rowNo < (_excelDataProvider.ExcelMaxRows + 1) ? rowNo.ToString() : string.Empty;
            if (!string.IsNullOrEmpty(retVal))
            {
                switch (_excelReferenceType)
                {
                    case ExcelReferenceType.AbsoluteRowAndColumn:
                    case ExcelReferenceType.AbsoluteRowRelativeColumn:
                        return "$" + retVal;
                    default:
                        return retVal;
                }
            }
            return retVal;
        }
    }
}
