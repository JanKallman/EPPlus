using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
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