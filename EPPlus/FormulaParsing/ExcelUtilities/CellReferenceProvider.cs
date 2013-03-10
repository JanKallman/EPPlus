using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class CellReferenceProvider
    {
        public virtual IEnumerable<string> GetReferencedAddresses(string cellFormula, ParsingContext context)
        {
            var resultCells = new List<string>();
            var r = context.Configuration.Lexer.Tokenize(cellFormula);
            var toAddresses = r.Where(x => x.TokenType == TokenType.ExcelAddress);
            foreach (var toAddress in toAddresses)
            {
                var rangeAddress = context.RangeAddressFactory.Create(toAddress.Value);
                var rangeCells = new List<string>();
                if (rangeAddress.FromRow < rangeAddress.ToRow || rangeAddress.FromCol < rangeAddress.ToCol)
                {
                    for (var col = rangeAddress.FromCol; col <= rangeAddress.ToCol; col++)
                    {
                        for (var row = rangeAddress.FromRow; row <= rangeAddress.ToRow; row++)
                        {
                            resultCells.Add(context.RangeAddressFactory.Create(col, row).Address);
                        }
                    }
                }
                else
                {
                    rangeCells.Add(toAddress.Value);
                }
                resultCells.AddRange(rangeCells);
            }
            return resultCells;
        }
    }
}
